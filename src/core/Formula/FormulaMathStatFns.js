export function createMathStatFns(ctx) {
    const {
        stack,
        gVal,
        gVals,
        gNums,
        execute,
        getFlatValues,
        buildCriteriaTester,
        ensureInteger,
        getNumericSorted,
        percentileInc,
        percentileExc,
        toDate
    } = ctx;


    const coerceToNumberA = (value) => {
        if (value === null || value === undefined || value === '') return null;
        if (value instanceof Error) throw value;
        if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
        if (typeof value === 'boolean') return value ? 1 : 0;
        const num = Number(value);
        return Number.isFinite(num) ? num : 0;
    };

    const collectNumbersA = (args) => {
        const values = [];
        gVals(args).forEach((value) => {
            const num = coerceToNumberA(value);
            if (num !== null) values.push(num);
        });
        return values;
    };

    const collectNumericPairs = (xArg, yArg) => {
        const xVals = flattenAny(xArg);
        const yVals = flattenAny(yArg);
        if (xVals.length !== yVals.length) throw new Error('#N/A');

        const pairs = [];
        for (let i = 0; i < xVals.length; i++) {
            const x = Number(xVals[i]);
            const y = Number(yVals[i]);
            if (Number.isFinite(x) && Number.isFinite(y)) pairs.push([x, y]);
        }
        return pairs;
    };

    const mean = (values) => values.reduce((sum, n) => sum + n, 0) / values.length;

    const varianceSample = (values) => {
        if (values.length < 2) throw new Error('#DIV/0!');
        const avg = mean(values);
        return values.reduce((sum, n) => sum + ((n - avg) ** 2), 0) / (values.length - 1);
    };

    const variancePopulation = (values) => {
        if (!values.length) throw new Error('#DIV/0!');
        const avg = mean(values);
        return values.reduce((sum, n) => sum + ((n - avg) ** 2), 0) / values.length;
    };

    const calcSlopeIntercept = (pairs) => {
        if (pairs.length < 2) throw new Error('#DIV/0!');
        const xAvg = mean(pairs.map((item) => item[0]));
        const yAvg = mean(pairs.map((item) => item[1]));

        let numerator = 0;
        let denominator = 0;
        pairs.forEach(([x, y]) => {
            numerator += (x - xAvg) * (y - yAvg);
            denominator += (x - xAvg) ** 2;
        });

        if (denominator === 0) throw new Error('#DIV/0!');
        const slope = numerator / denominator;
        return { slope, intercept: yAvg - slope * xAvg };
    };

    const EPS = 1e-12;

    const flattenAny = (arg) => {
        if (Array.isArray(arg)) return arg.flat(Infinity);
        const value = gVal(arg);
        if (Array.isArray(value)) return Array.isArray(value[0]) ? value.flat() : [...value];
        return [value];
    };

    const flattenNumeric = (arg) => flattenAny(arg)
        .map((v) => Number(v))
        .filter((v) => Number.isFinite(v));

    const toDateValue = (value) => {
        const date = toDate(value);
        if (!(date instanceof Date) || Number.isNaN(date.getTime())) throw new Error('#VALUE!');
        const out = new Date(date.getTime());
        out.setHours(0, 0, 0, 0);
        return out;
    };

    const dayCountDiff = (startRaw, endRaw) => {
        const start = toDateValue(startRaw);
        const end = toDateValue(endRaw);
        const diff = Math.floor((end.getTime() - start.getTime()) / 86400000);
        if (diff <= 0) throw new Error('#NUM!');
        return diff;
    };

    const basisYearDays = (basisRaw) => {
        const basis = basisRaw === undefined ? 0 : ensureInteger(Number(basisRaw), '#NUM!');
        if (basis < 0 || basis > 4) throw new Error('#NUM!');
        return (basis === 1 || basis === 3) ? 365 : 360;
    };

    const toSeries = (arg, fallbackLength = 0) => {
        if (arg === undefined) return Array.from({ length: fallbackLength }, (_, i) => i + 1);
        return flattenNumeric(arg);
    };

    const reshapeLikeArg = (arg, values) => {
        if (Array.isArray(arg)) {
            if (Array.isArray(arg[0])) {
                const rows = arg.length;
                const cols = arg[0].length;
                const out = [];
                let idx = 0;
                for (let r = 0; r < rows; r++) {
                    const row = [];
                    for (let c = 0; c < cols; c++) row.push(values[idx++]);
                    out.push(row);
                }
                return out;
            }
            return [values];
        }
        const raw = gVal(arg);
        if (!Array.isArray(raw)) return values[0];
        if (Array.isArray(raw[0])) {
            const rows = raw.length;
            const cols = raw[0].length;
            const out = [];
            let idx = 0;
            for (let r = 0; r < rows; r++) {
                const row = [];
                for (let c = 0; c < cols; c++) row.push(values[idx++]);
                out.push(row);
            }
            return out;
        }
        return [values];
    };


    const combination = (nValue, kValue) => {
        const n = ensureInteger(nValue, '#NUM!');
        const k = ensureInteger(kValue, '#NUM!');
        if (n < 0 || k < 0 || k > n) throw new Error('#NUM!');
        const kk = Math.min(k, n - k);
        let out = 1;
        for (let i = 1; i <= kk; i++) out = out * (n - kk + i) / i;
        return out;
    };

    const lnGamma = (zRaw) => {
        const z = Number(zRaw);
        if (!Number.isFinite(z) || z <= 0) throw new Error('#NUM!');
        const p = [
            676.5203681218851,
            -1259.1392167224028,
            771.3234287776531,
            -176.6150291621406,
            12.507343278686905,
            -0.13857109526572012,
            9.984369578019572e-6,
            1.5056327351493116e-7,
        ];
        if (z < 0.5) return Math.log(Math.PI) - Math.log(Math.sin(Math.PI * z)) - lnGamma(1 - z);
        let x = 0.9999999999998099;
        const tZ = z - 1;
        for (let i = 0; i < p.length; i++) x += p[i] / (tZ + i + 1);
        const t = tZ + p.length - 0.5;
        return 0.5 * Math.log(2 * Math.PI) + (tZ + 0.5) * Math.log(t) - t + Math.log(x);
    };

    const erf = (xRaw) => {
        const x = Number(xRaw);
        if (!Number.isFinite(x)) throw new Error('#VALUE!');
        const sign = x < 0 ? -1 : 1;
        const ax = Math.abs(x);
        const t = 1 / (1 + 0.3275911 * ax);
        const y = 1 - (((((1.061405429 * t - 1.453152027) * t + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t) * Math.exp(-(ax ** 2));
        return sign * y;
    };

    const normPdf = (x) => Math.exp(-(x ** 2) / 2) / Math.sqrt(2 * Math.PI);
    const normCdf = (x) => 0.5 * (1 + erf(x / Math.SQRT2));

    const normInv = (pRaw) => {
        const p = Number(pRaw);
        if (!Number.isFinite(p)) throw new Error('#VALUE!');
        if (p <= 0 || p >= 1) throw new Error('#NUM!');
        const a = [-39.69683028665376, 220.9460984245205, -275.9285104469687, 138.357751867269, -30.66479806614716, 2.506628277459239];
        const b = [-54.47609879822406, 161.5858368580409, -155.6989798598866, 66.80131188771972, -13.28068155288572];
        const c = [-0.007784894002430293, -0.3223964580411365, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
        const d = [0.007784695709041462, 0.3224671290700398, 2.445134137142996, 3.754408661907416];
        const plow = 0.02425;
        const phigh = 1 - plow;
        const poly = (coeffs, x) => coeffs.reduce((acc, cur) => acc * x + cur, 0);
        if (p < plow) {
            const q = Math.sqrt(-2 * Math.log(p));
            return poly(c, q) / poly([...d, 1], q);
        }
        if (p > phigh) {
            const q = Math.sqrt(-2 * Math.log(1 - p));
            return -(poly(c, q) / poly([...d, 1], q));
        }
        const q = p - 0.5;
        const r = q * q;
        return (poly(a, r) * q) / poly([...b, 1], r);
    };

    const regressionFromPairs = (pairs, forceZeroIntercept = false) => {
        if (pairs.length < 2) throw new Error('#DIV/0!');
        if (!forceZeroIntercept) return calcSlopeIntercept(pairs);
        const num = pairs.reduce((sum, [x, y]) => sum + x * y, 0);
        const den = pairs.reduce((sum, [x]) => sum + x * x, 0);
        if (den === 0) throw new Error('#DIV/0!');
        return { slope: num / den, intercept: 0 };
    };

    const binomPmf = (kRaw, nRaw, pRaw) => {
        const k = ensureInteger(kRaw, '#NUM!');
        const n = ensureInteger(nRaw, '#NUM!');
        const p = Number(pRaw);
        if (n < 0 || k < 0 || k > n || !Number.isFinite(p) || p < 0 || p > 1) throw new Error('#NUM!');
        return combination(n, k) * (p ** k) * ((1 - p) ** (n - k));
    };

    const negBinomPmf = (fRaw, sRaw, pRaw) => {
        const failures = ensureInteger(fRaw, '#NUM!');
        const successes = ensureInteger(sRaw, '#NUM!');
        const p = Number(pRaw);
        if (failures < 0 || successes <= 0 || !Number.isFinite(p) || p <= 0 || p > 1) throw new Error('#NUM!');
        return combination(failures + successes - 1, successes - 1) * ((1 - p) ** failures) * (p ** successes);
    };


    const financeType = (typeRaw) => {
        const type = typeRaw === undefined ? 0 : ensureInteger(Number(typeRaw), '#NUM!');
        if (type !== 0 && type !== 1) throw new Error('#NUM!');
        return type;
    };

    const pmtFrom = (rate, nper, pv, fv = 0, type = 0) => {
        if (nper === 0) throw new Error('#DIV/0!');
        if (Math.abs(rate) < EPS) return -(pv + fv) / nper;
        const pow = (1 + rate) ** nper;
        return -(rate * (fv + pv * pow)) / ((1 + rate * type) * (pow - 1));
    };

    const fvFrom = (rate, nper, pmt, pv, type = 0) => {
        if (Math.abs(rate) < EPS) return -(pv + pmt * nper);
        const pow = (1 + rate) ** nper;
        return -(pv * pow + pmt * (1 + rate * type) * (pow - 1) / rate);
    };

    const pvFrom = (rate, nper, pmt, fv, type = 0) => {
        if (Math.abs(rate) < EPS) return -(fv + pmt * nper);
        const pow = (1 + rate) ** nper;
        return -(fv + pmt * (1 + rate * type) * (pow - 1) / rate) / pow;
    };

    const nperFrom = (rate, pmt, pv, fv = 0, type = 0) => {
        if (Math.abs(rate) < EPS) {
            if (Math.abs(pmt) < EPS) throw new Error('#DIV/0!');
            return -(pv + fv) / pmt;
        }
        const term = pmt * (1 + rate * type) / rate;
        const numerator = term - fv;
        const denominator = pv + term;
        const ratio = numerator / denominator;
        if (!Number.isFinite(ratio) || ratio <= 0) throw new Error('#NUM!');
        return Math.log(ratio) / Math.log(1 + rate);
    };

    const rateFrom = (nper, pmt, pv, fv = 0, type = 0, guess = 0.1) => {
        const fn = (r) => {
            if (Math.abs(r) < EPS) return pv + pmt * nper + fv;
            const pow = (1 + r) ** nper;
            return pv * pow + pmt * (1 + r * type) * (pow - 1) / r + fv;
        };
        let r = Number.isFinite(guess) ? guess : 0.1;
        if (r <= -0.9999) r = -0.5;
        for (let i = 0; i < 100; i++) {
            const f0 = fn(r);
            if (Math.abs(f0) < 1e-10) return r;
            const h = 1e-6;
            const d = (fn(r + h) - fn(r - h)) / (2 * h);
            if (!Number.isFinite(d) || Math.abs(d) < 1e-12) break;
            const next = r - f0 / d;
            if (!Number.isFinite(next) || next <= -0.999999) break;
            r = next;
        }
        throw new Error('#NUM!');
    };

    const periodParts = (rate, nperRaw, pv, perRaw, fv = 0, type = 0) => {
        const nper = ensureInteger(nperRaw, '#NUM!');
        const per = ensureInteger(perRaw, '#NUM!');
        if (nper <= 0 || per < 1 || per > nper) throw new Error('#NUM!');
        const pmt = pmtFrom(rate, nper, pv, fv, type);
        if (Math.abs(rate) < EPS) return { pmt, ipmt: 0, ppmt: pmt };
        let balance = pv;
        for (let p = 1; p <= nper; p++) {
            if (type === 1) balance += pmt;
            const ipmt = -balance * rate;
            const ppmt = pmt - ipmt;
            if (p === per) return { pmt, ipmt, ppmt };
            balance += ppmt;
        }
        throw new Error('#NUM!');
    };

    const depreciationSchedule = (cost, salvage, life, factor = 2, noSwitch = false) => {
        const periods = Math.ceil(life);
        const out = [];
        let book = cost;
        for (let p = 1; p <= periods; p++) {
            if (book <= salvage + EPS) {
                out.push(0);
                continue;
            }
            const ddb = book * factor / life;
            const sl = (book - salvage) / Math.max(life - p + 1, EPS);
            let dep = noSwitch ? ddb : Math.max(ddb, sl);
            dep = Math.max(0, Math.min(dep, book - salvage));
            out.push(dep);
            book -= dep;
        }
        return out;
    };

    const calcCorrelation = (pairs) => {
        if (pairs.length < 2) throw new Error('#DIV/0!');
        const xAvg = mean(pairs.map((item) => item[0]));
        const yAvg = mean(pairs.map((item) => item[1]));

        let sumXY = 0;
        let sumXX = 0;
        let sumYY = 0;
        pairs.forEach(([x, y]) => {
            const dx = x - xAvg;
            const dy = y - yAvg;
            sumXY += dx * dy;
            sumXX += dx * dx;
            sumYY += dy * dy;
        });

        const denom = Math.sqrt(sumXX * sumYY);
        if (denom === 0) throw new Error('#DIV/0!');
        return sumXY / denom;
    };

    return {
        COUNTBLANK: () => getFlatValues(stack[0]).filter(v => v === '' || v === null || v === undefined).length,
        COUNTIFS: () => {
            if (stack.length < 2 || stack.length % 2 !== 0) throw new Error('#VALUE!');
            const firstRange = getFlatValues(stack[0]);
            const tests = [];
            for (let i = 0; i < stack.length; i += 2) {
                const range = getFlatValues(stack[i]);
                if (range.length !== firstRange.length) throw new Error('#VALUE!');
                tests.push({ range, test: buildCriteriaTester(gVal(stack[i + 1])) });
            }

            let count = 0;
            for (let i = 0; i < firstRange.length; i++) {
                if (tests.every(item => item.test(item.range[i]))) count++;
            }
            return count;
        },
        AVERAGEA: () => {
            const values = gVals(stack);
            let sum = 0;
            let count = 0;
            values.forEach(v => {
                if (v === null || v === undefined || v === '') return;
                if (typeof v === 'number' && Number.isFinite(v)) sum += v;
                else if (typeof v === 'boolean') sum += v ? 1 : 0;
                count++;
            });
            return count ? sum / count : 0;
        },

        AVEDEV: () => {
            const values = gNums(stack).filter(v => Number.isFinite(v));
            if (!values.length) throw new Error('#DIV/0!');
            const avg = mean(values);
            return values.reduce((sum, n) => sum + Math.abs(n - avg), 0) / values.length;
        },
        MAXA: () => {
            const values = gVals(stack);
            const nums = [];
            values.forEach(v => {
                if (v === null || v === undefined || v === '') return;
                if (typeof v === 'number' && Number.isFinite(v)) nums.push(v);
                else if (typeof v === 'boolean') nums.push(v ? 1 : 0);
                else nums.push(0);
            });
            return nums.length ? Math.max(...nums) : 0;
        },
        MINA: () => {
            const values = gVals(stack);
            const nums = [];
            values.forEach(v => {
                if (v === null || v === undefined || v === '') return;
                if (typeof v === 'number' && Number.isFinite(v)) nums.push(v);
                else if (typeof v === 'boolean') nums.push(v ? 1 : 0);
                else nums.push(0);
            });
            return nums.length ? Math.min(...nums) : 0;
        },
        AVERAGEIF: () => {
            const range = getFlatValues(stack[0]);
            const test = buildCriteriaTester(gVal(stack[1]));
            const avgRange = stack[2] !== undefined ? getFlatValues(stack[2]) : range;
            if (avgRange.length !== range.length) throw new Error('#VALUE!');
            let sum = 0;
            let count = 0;
            for (let i = 0; i < range.length; i++) {
                if (!test(range[i])) continue;
                const n = Number(avgRange[i]);
                if (Number.isFinite(n)) {
                    sum += n;
                    count++;
                }
            }
            if (!count) throw new Error('#DIV/0!');
            return sum / count;
        },
        AVERAGEIFS: () => {
            if (stack.length < 3 || (stack.length - 1) % 2 !== 0) throw new Error('#VALUE!');
            const avgRange = getFlatValues(stack[0]);
            const tests = [];
            for (let i = 1; i < stack.length; i += 2) {
                const criteriaRange = getFlatValues(stack[i]);
                if (criteriaRange.length !== avgRange.length) throw new Error('#VALUE!');
                tests.push({ range: criteriaRange, test: buildCriteriaTester(gVal(stack[i + 1])) });
            }

            let sum = 0;
            let count = 0;
            for (let i = 0; i < avgRange.length; i++) {
                if (!tests.every(item => item.test(item.range[i]))) continue;
                const n = Number(avgRange[i]);
                if (Number.isFinite(n)) {
                    sum += n;
                    count++;
                }
            }
            if (!count) throw new Error('#DIV/0!');
            return sum / count;
        },
        'VAR.S': () => execute('VAR', stack),
        'VAR.P': () => {
            const values = gNums(stack).filter(v => typeof v === 'number' && Number.isFinite(v));
            if (!values.length) return 0;
            const mean = values.reduce((a, b) => a + b, 0) / values.length;
            return values.reduce((sum, n) => sum + Math.pow(n - mean, 2), 0) / values.length;
        },
        'STDEV.S': () => execute('STDEV', stack),
        'STDEV.P': () => {
            const values = gNums(stack).filter(v => typeof v === 'number' && Number.isFinite(v));
            if (!values.length) return 0;
            const mean = values.reduce((a, b) => a + b, 0) / values.length;
            const variance = values.reduce((sum, n) => sum + Math.pow(n - mean, 2), 0) / values.length;
            return Math.sqrt(variance);
        },

        VARA: () => varianceSample(collectNumbersA(stack)),
        VARPA: () => variancePopulation(collectNumbersA(stack)),
        STDEVA: () => Math.sqrt(execute('VARA', stack)),
        STDEVPA: () => Math.sqrt(execute('VARPA', stack)),
        SUMIFS: () => {
            if (stack.length < 3 || (stack.length - 1) % 2 !== 0) throw new Error('#VALUE!');
            const sumRange = getFlatValues(stack[0]);
            const tests = [];
            for (let i = 1; i < stack.length; i += 2) {
                const criteriaRange = getFlatValues(stack[i]);
                if (criteriaRange.length !== sumRange.length) throw new Error('#VALUE!');
                tests.push({ range: criteriaRange, test: buildCriteriaTester(gVal(stack[i + 1])) });
            }
            let total = 0;
            for (let i = 0; i < sumRange.length; i++) {
                if (!tests.every(item => item.test(item.range[i]))) continue;
                const n = Number(sumRange[i]);
                if (Number.isFinite(n)) total += n;
            }
            return total;
        },
        COT: () => {
            const num = Number(gVal(stack[0]));
            if (!Number.isFinite(num) || num === 0) throw new Error('#DIV/0!');
            return 1 / Math.tan(num);
        },
        SEC: () => {
            const num = Number(gVal(stack[0]));
            if (!Number.isFinite(num) || Math.cos(num) === 0) throw new Error('#DIV/0!');
            return 1 / Math.cos(num);
        },
        CSC: () => {
            const num = Number(gVal(stack[0]));
            if (!Number.isFinite(num) || Math.sin(num) === 0) throw new Error('#DIV/0!');
            return 1 / Math.sin(num);
        },
        'CEILING.MATH': () => {
            const number = Number(gVal(stack[0]));
            const significanceRaw = stack[1] !== undefined ? Number(gVal(stack[1])) : 1;
            const significance = Math.abs(significanceRaw || 1);
            const mode = stack[2] !== undefined ? Number(gVal(stack[2])) : 0;
            if (!Number.isFinite(number) || !Number.isFinite(significance)) throw new Error('#VALUE!');
            if (number >= 0) return Math.ceil(number / significance) * significance;
            return mode === 0
                ? -Math.floor(Math.abs(number) / significance) * significance
                : -Math.ceil(Math.abs(number) / significance) * significance;
        },
        'FLOOR.MATH': () => {
            const number = Number(gVal(stack[0]));
            const significanceRaw = stack[1] !== undefined ? Number(gVal(stack[1])) : 1;
            const significance = Math.abs(significanceRaw || 1);
            const mode = stack[2] !== undefined ? Number(gVal(stack[2])) : 0;
            if (!Number.isFinite(number) || !Number.isFinite(significance)) throw new Error('#VALUE!');
            if (number >= 0) return Math.floor(number / significance) * significance;
            return mode === 0
                ? -Math.ceil(Math.abs(number) / significance) * significance
                : -Math.floor(Math.abs(number) / significance) * significance;
        },
        COMBIN: () => {
            const n = ensureInteger(gVal(stack[0]), '#NUM!');
            const k = ensureInteger(gVal(stack[1]), '#NUM!');
            if (n < 0 || k < 0 || k > n) throw new Error('#NUM!');
            const kk = Math.min(k, n - k);
            let out = 1;
            for (let i = 1; i <= kk; i++) out = out * (n - kk + i) / i;
            return out;
        },
        COMBINA: () => {
            const n = ensureInteger(gVal(stack[0]), '#NUM!');
            const k = ensureInteger(gVal(stack[1]), '#NUM!');
            if (n < 0 || k < 0) throw new Error('#NUM!');
            return execute('COMBIN', [n + k - 1, k]);
        },

        PERMUT: () => {
            const n = ensureInteger(gVal(stack[0]), '#NUM!');
            const k = ensureInteger(gVal(stack[1]), '#NUM!');
            if (n < 0 || k < 0 || k > n) throw new Error('#NUM!');
            let out = 1;
            for (let i = 0; i < k; i++) out *= (n - i);
            return out;
        },
        PERMUTATIONA: () => {
            const n = ensureInteger(gVal(stack[0]), '#NUM!');
            const k = ensureInteger(gVal(stack[1]), '#NUM!');
            if (n < 0 || k < 0) throw new Error('#NUM!');
            return n ** k;
        },
        MULTINOMIAL: () => {
            const nums = gNums(stack).filter(v => typeof v === 'number' && Number.isFinite(v));
            if (!nums.length || nums.some(v => v < 0 || !Number.isInteger(v))) throw new Error('#NUM!');
            const fact = (n) => {
                let out = 1;
                for (let i = 2; i <= n; i++) out *= i;
                return out;
            };
            const total = nums.reduce((a, b) => a + b, 0);
            return fact(total) / nums.reduce((acc, n) => acc * fact(n), 1);
        },
        PERCENTOF: () => {
            const value = Number(gVal(stack[0]));
            const total = Number(gVal(stack[1]));
            if (!Number.isFinite(value) || !Number.isFinite(total)) throw new Error('#VALUE!');
            if (total === 0) throw new Error('#DIV/0!');
            return value / total;
        },

        DEVSQ: () => {
            const values = gNums(stack).filter(v => Number.isFinite(v));
            if (!values.length) throw new Error('#DIV/0!');
            const avg = mean(values);
            return values.reduce((sum, n) => sum + ((n - avg) ** 2), 0);
        },
        GEOMEAN: () => {
            const values = gNums(stack).filter(v => Number.isFinite(v));
            if (!values.length) throw new Error('#DIV/0!');
            if (values.some(v => v <= 0)) throw new Error('#NUM!');
            const logSum = values.reduce((sum, n) => sum + Math.log(n), 0);
            return Math.exp(logSum / values.length);
        },
        HARMEAN: () => {
            const values = gNums(stack).filter(v => Number.isFinite(v));
            if (!values.length) throw new Error('#DIV/0!');
            if (values.some(v => v <= 0)) throw new Error('#NUM!');
            const invSum = values.reduce((sum, n) => sum + (1 / n), 0);
            return values.length / invSum;
        },
        PHI: () => {
            const x = Number(gVal(stack[0]));
            if (!Number.isFinite(x)) throw new Error('#VALUE!');
            return Math.exp(-(x ** 2) / 2) / Math.sqrt(2 * Math.PI);
        },
        STANDARDIZE: () => {
            const x = Number(gVal(stack[0]));
            const avg = Number(gVal(stack[1]));
            const stdev = Number(gVal(stack[2]));
            if (!Number.isFinite(x) || !Number.isFinite(avg) || !Number.isFinite(stdev)) throw new Error('#VALUE!');
            if (stdev <= 0) throw new Error('#NUM!');
            return (x - avg) / stdev;
        },
        FISHER: () => {
            const x = Number(gVal(stack[0]));
            if (!Number.isFinite(x)) throw new Error('#VALUE!');
            if (x <= -1 || x >= 1) throw new Error('#NUM!');
            return 0.5 * Math.log((1 + x) / (1 - x));
        },
        FISHERINV: () => {
            const y = Number(gVal(stack[0]));
            if (!Number.isFinite(y)) throw new Error('#VALUE!');
            const exp2y = Math.exp(2 * y);
            return (exp2y - 1) / (exp2y + 1);
        },
        CORREL: () => calcCorrelation(collectNumericPairs(stack[0], stack[1])),
        PEARSON: () => execute('CORREL', stack),
        RSQ: () => {
            const corr = execute('CORREL', stack);
            return corr * corr;
        },
        SLOPE: () => calcSlopeIntercept(collectNumericPairs(stack[1], stack[0])).slope,
        INTERCEPT: () => calcSlopeIntercept(collectNumericPairs(stack[1], stack[0])).intercept,
        COTH: () => {
            const num = Number(gVal(stack[0]));
            if (!Number.isFinite(num) || num === 0) throw new Error('#DIV/0!');
            return Math.cosh(num) / Math.sinh(num);
        },
        CSCH: () => {
            const num = Number(gVal(stack[0]));
            if (!Number.isFinite(num) || Math.sinh(num) === 0) throw new Error('#DIV/0!');
            return 1 / Math.sinh(num);
        },
        SECH: () => {
            const num = Number(gVal(stack[0]));
            if (!Number.isFinite(num)) throw new Error('#VALUE!');
            return 1 / Math.cosh(num);
        },
        BASE: () => {
            const number = ensureInteger(gVal(stack[0]), '#NUM!');
            const radix = ensureInteger(gVal(stack[1]), '#NUM!');
            const minLength = stack[2] !== undefined ? ensureInteger(gVal(stack[2]), '#NUM!') : 0;
            if (number < 0 || radix < 2 || radix > 36 || minLength < 0) throw new Error('#NUM!');
            const out = number.toString(radix).toUpperCase();
            return minLength > out.length ? `${'0'.repeat(minLength - out.length)}${out}` : out;
        },
        DECIMAL: () => {
            const text = String(gVal(stack[0]) ?? '').trim();
            const radix = ensureInteger(gVal(stack[1]), '#NUM!');
            if (!text || radix < 2 || radix > 36) throw new Error('#NUM!');
            const out = parseInt(text, radix);
            if (!Number.isFinite(out)) throw new Error('#NUM!');
            return out;
        },
        RANK: () => execute('RANK.EQ', stack),
        'RANK.EQ': () => {
            const number = Number(gVal(stack[0]));
            const values = getNumericSorted(stack[1]);
            const order = stack[2] !== undefined ? Number(gVal(stack[2])) : 0;
            if (!Number.isFinite(number) || !values.length) throw new Error('#N/A');
            const ranked = order === 1 ? [...values] : [...values].reverse();
            const idx = ranked.findIndex(v => v === number);
            if (idx < 0) throw new Error('#N/A');
            return idx + 1;
        },
        MEDIAN: () => {
            const sorted = getNumericSorted(gVals(stack));
            if (!sorted.length) return 0;
            const mid = Math.floor(sorted.length / 2);
            return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
        },
        LARGE: () => {
            const sorted = getNumericSorted(stack[0]).reverse();
            const k = ensureInteger(gVal(stack[1]), '#NUM!');
            if (k < 1 || k > sorted.length) throw new Error('#NUM!');
            return sorted[k - 1];
        },
        SMALL: () => {
            const sorted = getNumericSorted(stack[0]);
            const k = ensureInteger(gVal(stack[1]), '#NUM!');
            if (k < 1 || k > sorted.length) throw new Error('#NUM!');
            return sorted[k - 1];
        },
        MODE: () => execute('MODE.SNGL', stack),
        'MODE.SNGL': () => {
            const values = getNumericSorted(gVals(stack));
            if (!values.length) throw new Error('#N/A');
            const counts = new Map();
            values.forEach(v => counts.set(v, (counts.get(v) || 0) + 1));
            let best = null;
            let bestCount = 0;
            counts.forEach((count, value) => {
                if (count > bestCount || (count === bestCount && best !== null && value < best)) {
                    best = value;
                    bestCount = count;
                }
            });
            if (bestCount <= 1) throw new Error('#N/A');
            return best;
        },
        PERCENTILE: () => execute('PERCENTILE.INC', stack),
        'PERCENTILE.INC': () => percentileInc(getNumericSorted(stack[0]), Number(gVal(stack[1]))),
        'PERCENTILE.EXC': () => percentileExc(getNumericSorted(stack[0]), Number(gVal(stack[1]))),
        QUARTILE: () => execute('QUARTILE.INC', stack),
        'QUARTILE.INC': () => {
            const sorted = getNumericSorted(stack[0]);
            const quart = ensureInteger(gVal(stack[1]), '#NUM!');
            if (quart < 0 || quart > 4) throw new Error('#NUM!');
            return percentileInc(sorted, quart / 4);
        },
        'QUARTILE.EXC': () => {
            const sorted = getNumericSorted(stack[0]);
            const quart = ensureInteger(gVal(stack[1]), '#NUM!');
            if (quart < 1 || quart > 3) throw new Error('#NUM!');
            return percentileExc(sorted, quart / 4);
        },
        PERCENTRANK: () => execute('PERCENTRANK.INC', stack),
        'PERCENTRANK.INC': () => {
            const sorted = getNumericSorted(stack[0]);
            const x = Number(gVal(stack[1]));
            const sig = stack[2] !== undefined ? ensureInteger(gVal(stack[2]), '#NUM!') : 3;
            if (!sorted.length || !Number.isFinite(x)) throw new Error('#N/A');
            if (x < sorted[0] || x > sorted[sorted.length - 1]) throw new Error('#N/A');
            if (sorted.length === 1) return 1;

            let rank = 0;
            for (let i = 0; i < sorted.length; i++) {
                if (sorted[i] === x) {
                    rank = i / (sorted.length - 1);
                    break;
                }
                if (sorted[i] < x && i + 1 < sorted.length && sorted[i + 1] > x) {
                    const ratio = (x - sorted[i]) / (sorted[i + 1] - sorted[i]);
                    rank = (i + ratio) / (sorted.length - 1);
                    break;
                }
            }
            const factor = 10 ** sig;
            return Math.round(rank * factor) / factor;
        },
        'PERCENTRANK.EXC': () => {
            const sorted = getNumericSorted(stack[0]);
            const x = Number(gVal(stack[1]));
            const sig = stack[2] !== undefined ? ensureInteger(gVal(stack[2]), '#NUM!') : 3;
            if (!sorted.length || !Number.isFinite(x)) throw new Error('#N/A');
            if (x <= sorted[0] || x >= sorted[sorted.length - 1]) throw new Error('#N/A');

            let rank = 0;
            for (let i = 0; i < sorted.length; i++) {
                if (sorted[i] === x) {
                    rank = (i + 1) / (sorted.length + 1);
                    break;
                }
                if (sorted[i] < x && i + 1 < sorted.length && sorted[i + 1] > x) {
                    const ratio = (x - sorted[i]) / (sorted[i + 1] - sorted[i]);
                    rank = (i + 1 + ratio) / (sorted.length + 1);
                    break;
                }
            }
            const factor = 10 ** sig;
            return Math.round(rank * factor) / factor;
        },

        MAXIFS: () => {
            if (stack.length < 3 || (stack.length - 1) % 2 !== 0) throw new Error('#VALUE!');
            const maxRange = getFlatValues(stack[0]);
            const tests = [];
            for (let i = 1; i < stack.length; i += 2) {
                const criteriaRange = getFlatValues(stack[i]);
                if (criteriaRange.length !== maxRange.length) throw new Error('#VALUE!');
                tests.push({ range: criteriaRange, test: buildCriteriaTester(gVal(stack[i + 1])) });
            }
            const nums = [];
            for (let i = 0; i < maxRange.length; i++) {
                if (!tests.every((item) => item.test(item.range[i]))) continue;
                const num = Number(maxRange[i]);
                if (Number.isFinite(num)) nums.push(num);
            }
            return nums.length ? Math.max(...nums) : 0;
        },
        MINIFS: () => {
            if (stack.length < 3 || (stack.length - 1) % 2 !== 0) throw new Error('#VALUE!');
            const minRange = getFlatValues(stack[0]);
            const tests = [];
            for (let i = 1; i < stack.length; i += 2) {
                const criteriaRange = getFlatValues(stack[i]);
                if (criteriaRange.length !== minRange.length) throw new Error('#VALUE!');
                tests.push({ range: criteriaRange, test: buildCriteriaTester(gVal(stack[i + 1])) });
            }
            const nums = [];
            for (let i = 0; i < minRange.length; i++) {
                if (!tests.every((item) => item.test(item.range[i]))) continue;
                const num = Number(minRange[i]);
                if (Number.isFinite(num)) nums.push(num);
            }
            return nums.length ? Math.min(...nums) : 0;
        },
        'MODE.MULT': () => {
            const values = gNums(stack).filter((v) => Number.isFinite(v));
            if (!values.length) throw new Error('#N/A');
            const countMap = new Map();
            values.forEach((v) => countMap.set(v, (countMap.get(v) || 0) + 1));
            let best = 0;
            countMap.forEach((count) => { if (count > best) best = count; });
            if (best <= 1) throw new Error('#N/A');
            const modes = [...countMap.entries()].filter(([, count]) => count === best).map(([value]) => value).sort((a, b) => a - b);
            return modes.map((value) => [value]);
        },
        COVAR: () => execute('COVARIANCE.P', stack),
        'COVARIANCE.P': () => {
            const pairs = collectNumericPairs(stack[0], stack[1]);
            if (!pairs.length) throw new Error('#DIV/0!');
            const xAvg = mean(pairs.map((item) => item[0]));
            const yAvg = mean(pairs.map((item) => item[1]));
            return pairs.reduce((sum, [x, y]) => sum + (x - xAvg) * (y - yAvg), 0) / pairs.length;
        },
        'COVARIANCE.S': () => {
            const pairs = collectNumericPairs(stack[0], stack[1]);
            if (pairs.length < 2) throw new Error('#DIV/0!');
            const xAvg = mean(pairs.map((item) => item[0]));
            const yAvg = mean(pairs.map((item) => item[1]));
            return pairs.reduce((sum, [x, y]) => sum + (x - xAvg) * (y - yAvg), 0) / (pairs.length - 1);
        },
        DOLLARDE: () => {
            const dollar = Number(gVal(stack[0]));
            const fraction = ensureInteger(gVal(stack[1]), '#NUM!');
            if (!Number.isFinite(dollar) || fraction <= 0) throw new Error('#NUM!');
            const intPart = Math.trunc(dollar);
            const fracDigits = Math.abs(dollar - intPart);
            return intPart + (fracDigits * 100) / fraction;
        },
        DOLLARFR: () => {
            const decimal = Number(gVal(stack[0]));
            const fraction = ensureInteger(gVal(stack[1]), '#NUM!');
            if (!Number.isFinite(decimal) || fraction <= 0) throw new Error('#NUM!');
            const sign = decimal < 0 ? -1 : 1;
            const abs = Math.abs(decimal);
            const intPart = Math.trunc(abs);
            const frac = abs - intPart;
            return sign * (intPart + (frac * fraction) / 100);
        },
        EFFECT: () => {
            const nominal = Number(gVal(stack[0]));
            const npery = ensureInteger(gVal(stack[1]), '#NUM!');
            if (!Number.isFinite(nominal) || npery < 1) throw new Error('#NUM!');
            return (1 + nominal / npery) ** npery - 1;
        },
        NOMINAL: () => {
            const effect = Number(gVal(stack[0]));
            const npery = ensureInteger(gVal(stack[1]), '#NUM!');
            if (!Number.isFinite(effect) || effect <= -1 || npery < 1) throw new Error('#NUM!');
            return npery * (((1 + effect) ** (1 / npery)) - 1);
        },


        PV: () => {
            const rate = Number(gVal(stack[0]));
            const nper = Number(gVal(stack[1]));
            const pmt = Number(gVal(stack[2]));
            const fv = stack[3] !== undefined ? Number(gVal(stack[3])) : 0;
            const type = financeType(stack[4] !== undefined ? gVal(stack[4]) : 0);
            if (![rate, nper, pmt, fv].every(Number.isFinite)) throw new Error('#VALUE!');
            if (nper <= 0) throw new Error('#NUM!');
            return pvFrom(rate, nper, pmt, fv, type);
        },
        FV: () => {
            const rate = Number(gVal(stack[0]));
            const nper = Number(gVal(stack[1]));
            const pmt = Number(gVal(stack[2]));
            const pv = stack[3] !== undefined ? Number(gVal(stack[3])) : 0;
            const type = financeType(stack[4] !== undefined ? gVal(stack[4]) : 0);
            if (![rate, nper, pmt, pv].every(Number.isFinite)) throw new Error('#VALUE!');
            if (nper <= 0) throw new Error('#NUM!');
            return fvFrom(rate, nper, pmt, pv, type);
        },
        PMT: () => {
            const rate = Number(gVal(stack[0]));
            const nper = Number(gVal(stack[1]));
            const pv = Number(gVal(stack[2]));
            const fv = stack[3] !== undefined ? Number(gVal(stack[3])) : 0;
            const type = financeType(stack[4] !== undefined ? gVal(stack[4]) : 0);
            if (![rate, nper, pv, fv].every(Number.isFinite)) throw new Error('#VALUE!');
            if (nper <= 0) throw new Error('#NUM!');
            return pmtFrom(rate, nper, pv, fv, type);
        },
        NPER: () => {
            const rate = Number(gVal(stack[0]));
            const pmt = Number(gVal(stack[1]));
            const pv = Number(gVal(stack[2]));
            const fv = stack[3] !== undefined ? Number(gVal(stack[3])) : 0;
            const type = financeType(stack[4] !== undefined ? gVal(stack[4]) : 0);
            if (![rate, pmt, pv, fv].every(Number.isFinite)) throw new Error('#VALUE!');
            return nperFrom(rate, pmt, pv, fv, type);
        },
        RATE: () => {
            const nper = Number(gVal(stack[0]));
            const pmt = Number(gVal(stack[1]));
            const pv = Number(gVal(stack[2]));
            const fv = stack[3] !== undefined ? Number(gVal(stack[3])) : 0;
            const type = financeType(stack[4] !== undefined ? gVal(stack[4]) : 0);
            const guess = stack[5] !== undefined ? Number(gVal(stack[5])) : 0.1;
            if (![nper, pmt, pv, fv, guess].every(Number.isFinite)) throw new Error('#VALUE!');
            if (nper <= 0) throw new Error('#NUM!');
            return rateFrom(nper, pmt, pv, fv, type, guess);
        },
        IPMT: () => {
            const rate = Number(gVal(stack[0]));
            const per = Number(gVal(stack[1]));
            const nper = Number(gVal(stack[2]));
            const pv = Number(gVal(stack[3]));
            const fv = stack[4] !== undefined ? Number(gVal(stack[4])) : 0;
            const type = financeType(stack[5] !== undefined ? gVal(stack[5]) : 0);
            if (![rate, per, nper, pv, fv].every(Number.isFinite)) throw new Error('#VALUE!');
            return periodParts(rate, nper, pv, per, fv, type).ipmt;
        },
        PPMT: () => {
            const rate = Number(gVal(stack[0]));
            const per = Number(gVal(stack[1]));
            const nper = Number(gVal(stack[2]));
            const pv = Number(gVal(stack[3]));
            const fv = stack[4] !== undefined ? Number(gVal(stack[4])) : 0;
            const type = financeType(stack[5] !== undefined ? gVal(stack[5]) : 0);
            if (![rate, per, nper, pv, fv].every(Number.isFinite)) throw new Error('#VALUE!');
            return periodParts(rate, nper, pv, per, fv, type).ppmt;
        },
        CUMIPMT: () => {
            const rate = Number(gVal(stack[0]));
            const nper = Number(gVal(stack[1]));
            const pv = Number(gVal(stack[2]));
            const start = ensureInteger(gVal(stack[3]), '#NUM!');
            const end = ensureInteger(gVal(stack[4]), '#NUM!');
            const type = financeType(gVal(stack[5]));
            if (![rate, nper, pv].every(Number.isFinite) || rate <= 0 || nper <= 0 || pv <= 0) throw new Error('#NUM!');
            if (start < 1 || end < start) throw new Error('#NUM!');
            let total = 0;
            for (let per = start; per <= end; per++) total += periodParts(rate, nper, pv, per, 0, type).ipmt;
            return total;
        },
        CUMPRINC: () => {
            const rate = Number(gVal(stack[0]));
            const nper = Number(gVal(stack[1]));
            const pv = Number(gVal(stack[2]));
            const start = ensureInteger(gVal(stack[3]), '#NUM!');
            const end = ensureInteger(gVal(stack[4]), '#NUM!');
            const type = financeType(gVal(stack[5]));
            if (![rate, nper, pv].every(Number.isFinite) || rate <= 0 || nper <= 0 || pv <= 0) throw new Error('#NUM!');
            if (start < 1 || end < start) throw new Error('#NUM!');
            let total = 0;
            for (let per = start; per <= end; per++) total += periodParts(rate, nper, pv, per, 0, type).ppmt;
            return total;
        },


        ISPMT: () => {
            const rate = Number(gVal(stack[0]));
            const per = Number(gVal(stack[1]));
            const nper = Number(gVal(stack[2]));
            const pv = Number(gVal(stack[3]));
            if (![rate, per, nper, pv].every(Number.isFinite) || nper <= 0) throw new Error('#NUM!');
            return pv * rate * (per / nper - 1);
        },
        NPV: () => {
            const rate = Number(gVal(stack[0]));
            if (!Number.isFinite(rate) || rate <= -1) throw new Error('#NUM!');
            const values = gVals(stack.slice(1)).map((v) => Number(v)).filter((v) => Number.isFinite(v));
            return values.reduce((sum, cash, i) => sum + cash / ((1 + rate) ** (i + 1)), 0);
        },
        IRR: () => {
            const values = gVals([stack[0]]).map((v) => Number(v)).filter((v) => Number.isFinite(v));
            const guess = stack[1] !== undefined ? Number(gVal(stack[1])) : 0.1;
            if (!values.length || !values.some((v) => v > 0) || !values.some((v) => v < 0)) throw new Error('#NUM!');
            const fn = (r) => values.reduce((sum, cash, i) => sum + cash / ((1 + r) ** i), 0);
            let r = Number.isFinite(guess) ? guess : 0.1;
            if (r <= -0.9999) r = -0.5;
            for (let i = 0; i < 100; i++) {
                const f0 = fn(r);
                if (Math.abs(f0) < 1e-10) return r;
                const h = 1e-6;
                const d = (fn(r + h) - fn(r - h)) / (2 * h);
                if (!Number.isFinite(d) || Math.abs(d) < 1e-12) break;
                const next = r - f0 / d;
                if (!Number.isFinite(next) || next <= -0.999999) break;
                r = next;
            }
            throw new Error('#NUM!');
        },
        XNPV: () => {
            const rate = Number(gVal(stack[0]));
            if (!Number.isFinite(rate) || rate <= -1) throw new Error('#NUM!');
            const values = flattenAny(stack[1]).map((v) => Number(v));
            const dates = flattenAny(stack[2]).map((d) => toDateValue(d));
            if (!values.length || values.length !== dates.length) throw new Error('#NUM!');
            const base = dates[0].getTime();
            return values.reduce((sum, cash, i) => {
                if (!Number.isFinite(cash)) return sum;
                const t = (dates[i].getTime() - base) / 86400000 / 365;
                return sum + cash / ((1 + rate) ** t);
            }, 0);
        },
        XIRR: () => {
            const values = flattenAny(stack[0]).map((v) => Number(v));
            const dates = flattenAny(stack[1]).map((d) => toDateValue(d));
            const guess = stack[2] !== undefined ? Number(gVal(stack[2])) : 0.1;
            if (!values.length || values.length !== dates.length) throw new Error('#NUM!');
            if (!values.some((v) => v > 0) || !values.some((v) => v < 0)) throw new Error('#NUM!');
            const base = dates[0].getTime();
            const fn = (r) => values.reduce((sum, cash, i) => {
                const t = (dates[i].getTime() - base) / 86400000 / 365;
                return sum + cash / ((1 + r) ** t);
            }, 0);
            let r = Number.isFinite(guess) ? guess : 0.1;
            if (r <= -0.9999) r = -0.5;
            for (let i = 0; i < 120; i++) {
                const f0 = fn(r);
                if (Math.abs(f0) < 1e-10) return r;
                const h = 1e-6;
                const d = (fn(r + h) - fn(r - h)) / (2 * h);
                if (!Number.isFinite(d) || Math.abs(d) < 1e-12) break;
                const next = r - f0 / d;
                if (!Number.isFinite(next) || next <= -0.999999) break;
                r = next;
            }
            throw new Error('#NUM!');
        },
        MIRR: () => {
            const values = gVals([stack[0]]).map((v) => Number(v)).filter((v) => Number.isFinite(v));
            const financeRate = Number(gVal(stack[1]));
            const reinvestRate = Number(gVal(stack[2]));
            if (!values.length || !Number.isFinite(financeRate) || !Number.isFinite(reinvestRate)) throw new Error('#VALUE!');
            if (!values.some((v) => v > 0) || !values.some((v) => v < 0)) throw new Error('#DIV/0!');
            const n = values.length;
            let pvNeg = 0;
            let fvPos = 0;
            values.forEach((cash, i) => {
                if (cash < 0) pvNeg += cash / ((1 + financeRate) ** i);
                if (cash > 0) fvPos += cash * ((1 + reinvestRate) ** (n - 1 - i));
            });
            if (pvNeg === 0 || fvPos <= 0) throw new Error('#DIV/0!');
            return ((-fvPos / pvNeg) ** (1 / (n - 1))) - 1;
        },
        FVSCHEDULE: () => {
            const principal = Number(gVal(stack[0]));
            if (!Number.isFinite(principal)) throw new Error('#VALUE!');
            const schedule = flattenNumeric(stack[1]);
            return schedule.reduce((acc, rate) => acc * (1 + rate), principal);
        },
        PDURATION: () => {
            const rate = Number(gVal(stack[0]));
            const pv = Number(gVal(stack[1]));
            const fv = Number(gVal(stack[2]));
            if (![rate, pv, fv].every(Number.isFinite) || rate <= 0 || pv <= 0 || fv <= 0) throw new Error('#NUM!');
            return Math.log(fv / pv) / Math.log(1 + rate);
        },
        RRI: () => {
            const nper = Number(gVal(stack[0]));
            const pv = Number(gVal(stack[1]));
            const fv = Number(gVal(stack[2]));
            if (![nper, pv, fv].every(Number.isFinite) || nper <= 0 || pv === 0) throw new Error('#NUM!');
            if ((pv > 0 && fv < 0) || (pv < 0 && fv > 0)) throw new Error('#NUM!');
            return (fv / pv) ** (1 / nper) - 1;
        },
        SLN: () => {
            const cost = Number(gVal(stack[0]));
            const salvage = Number(gVal(stack[1]));
            const life = Number(gVal(stack[2]));
            if (![cost, salvage, life].every(Number.isFinite) || life <= 0) throw new Error('#NUM!');
            return (cost - salvage) / life;
        },
        SYD: () => {
            const cost = Number(gVal(stack[0]));
            const salvage = Number(gVal(stack[1]));
            const life = Number(gVal(stack[2]));
            const period = Number(gVal(stack[3]));
            if (![cost, salvage, life, period].every(Number.isFinite) || life <= 0 || period <= 0 || period > life) throw new Error('#NUM!');
            return (cost - salvage) * (life - period + 1) * 2 / (life * (life + 1));
        },


        DDB: () => {
            const cost = Number(gVal(stack[0]));
            const salvage = Number(gVal(stack[1]));
            const life = Number(gVal(stack[2]));
            const period = ensureInteger(gVal(stack[3]), '#NUM!');
            const factor = stack[4] !== undefined ? Number(gVal(stack[4])) : 2;
            if (![cost, salvage, life, factor].every(Number.isFinite) || cost < 0 || salvage < 0 || life <= 0 || factor <= 0 || period < 1) throw new Error('#NUM!');
            let book = cost;
            let dep = 0;
            for (let p = 1; p <= period; p++) {
                dep = Math.min(book * factor / life, book - salvage);
                dep = Math.max(0, dep);
                book -= dep;
            }
            return dep;
        },
        DB: () => {
            const cost = Number(gVal(stack[0]));
            const salvage = Number(gVal(stack[1]));
            const life = Number(gVal(stack[2]));
            const period = ensureInteger(gVal(stack[3]), '#NUM!');
            const month = stack[4] !== undefined ? ensureInteger(gVal(stack[4]), '#NUM!') : 12;
            if (![cost, salvage, life].every(Number.isFinite) || cost <= 0 || salvage < 0 || life <= 0 || period < 1 || month < 1 || month > 12) throw new Error('#NUM!');
            const rate = Math.round((1 - (salvage / cost) ** (1 / life)) * 1000) / 1000;
            if (period === 1) return cost * rate * month / 12;
            let book = cost - cost * rate * month / 12;
            let dep = 0;
            const maxPeriod = Math.floor(life) + 1;
            if (period > maxPeriod) return 0;
            for (let p = 2; p <= period; p++) {
                dep = p === maxPeriod ? book * rate * (12 - month) / 12 : book * rate;
                book -= dep;
            }
            return dep;
        },
        VDB: () => {
            const cost = Number(gVal(stack[0]));
            const salvage = Number(gVal(stack[1]));
            const life = Number(gVal(stack[2]));
            const start = Number(gVal(stack[3]));
            const end = Number(gVal(stack[4]));
            const factor = stack[5] !== undefined ? Number(gVal(stack[5])) : 2;
            const noSwitch = stack[6] !== undefined ? !!gVal(stack[6]) : false;
            if (![cost, salvage, life, start, end, factor].every(Number.isFinite) || cost < 0 || salvage < 0 || life <= 0 || factor <= 0 || start < 0 || end <= start) throw new Error('#NUM!');
            const deps = depreciationSchedule(cost, salvage, life, factor, noSwitch);
            let total = 0;
            for (let i = 0; i < deps.length; i++) {
                const pStart = i;
                const pEnd = i + 1;
                const overlap = Math.max(0, Math.min(end, pEnd) - Math.max(start, pStart));
                if (overlap > 0) total += deps[i] * overlap;
            }
            return total;
        },
        INTRATE: () => {
            const settlement = gVal(stack[0]);
            const maturity = gVal(stack[1]);
            const investment = Number(gVal(stack[2]));
            const redemption = Number(gVal(stack[3]));
            const basis = stack[4] !== undefined ? gVal(stack[4]) : 0;
            if (![investment, redemption].every(Number.isFinite) || investment <= 0 || redemption <= 0) throw new Error('#NUM!');
            const days = dayCountDiff(settlement, maturity);
            const yearDays = basisYearDays(basis);
            return (redemption - investment) / investment * (yearDays / days);
        },
        DISC: () => {
            const settlement = gVal(stack[0]);
            const maturity = gVal(stack[1]);
            const pr = Number(gVal(stack[2]));
            const redemption = Number(gVal(stack[3]));
            const basis = stack[4] !== undefined ? gVal(stack[4]) : 0;
            if (![pr, redemption].every(Number.isFinite) || pr <= 0 || redemption <= 0) throw new Error('#NUM!');
            const days = dayCountDiff(settlement, maturity);
            const yearDays = basisYearDays(basis);
            return (redemption - pr) / redemption * (yearDays / days);
        },
        RECEIVED: () => {
            const settlement = gVal(stack[0]);
            const maturity = gVal(stack[1]);
            const investment = Number(gVal(stack[2]));
            const discount = Number(gVal(stack[3]));
            const basis = stack[4] !== undefined ? gVal(stack[4]) : 0;
            if (![investment, discount].every(Number.isFinite) || investment <= 0 || discount <= 0) throw new Error('#NUM!');
            const days = dayCountDiff(settlement, maturity);
            const yearDays = basisYearDays(basis);
            const denom = 1 - discount * days / yearDays;
            if (denom <= 0) throw new Error('#NUM!');
            return investment / denom;
        },
        GAMMALN: () => lnGamma(gVal(stack[0])),
        'GAMMALN.PRECISE': () => execute('GAMMALN', stack),
        GAMMA: () => Math.exp(lnGamma(gVal(stack[0]))),


        'NORM.S.DIST': () => {
            const z = Number(gVal(stack[0]));
            const cumulative = stack[1] !== undefined ? !!gVal(stack[1]) : true;
            if (!Number.isFinite(z)) throw new Error('#VALUE!');
            return cumulative ? normCdf(z) : normPdf(z);
        },
        NORMSDIST: () => execute('NORM.S.DIST', [stack[0], true]),
        'NORM.S.INV': () => normInv(gVal(stack[0])),
        NORMSINV: () => execute('NORM.S.INV', stack),
        'NORM.DIST': () => {
            const x = Number(gVal(stack[0]));
            const mu = Number(gVal(stack[1]));
            const sigma = Number(gVal(stack[2]));
            const cumulative = !!gVal(stack[3]);
            if (![x, mu, sigma].every(Number.isFinite)) throw new Error('#VALUE!');
            if (sigma <= 0) throw new Error('#NUM!');
            const z = (x - mu) / sigma;
            return cumulative ? normCdf(z) : normPdf(z) / sigma;
        },
        NORMDIST: () => execute('NORM.DIST', [stack[0], stack[1], stack[2], true]),
        'NORM.INV': () => {
            const p = Number(gVal(stack[0]));
            const mu = Number(gVal(stack[1]));
            const sigma = Number(gVal(stack[2]));
            if (![p, mu, sigma].every(Number.isFinite)) throw new Error('#VALUE!');
            if (sigma <= 0) throw new Error('#NUM!');
            return mu + sigma * normInv(p);
        },
        NORMINV: () => execute('NORM.INV', stack),
        'LOGNORM.DIST': () => {
            const x = Number(gVal(stack[0]));
            const mu = Number(gVal(stack[1]));
            const sigma = Number(gVal(stack[2]));
            const cumulative = !!gVal(stack[3]);
            if (![x, mu, sigma].every(Number.isFinite)) throw new Error('#VALUE!');
            if (x <= 0 || sigma <= 0) throw new Error('#NUM!');
            const z = (Math.log(x) - mu) / sigma;
            if (cumulative) return normCdf(z);
            return Math.exp(-(z ** 2) / 2) / (x * sigma * Math.sqrt(2 * Math.PI));
        },
        LOGNORMDIST: () => execute('LOGNORM.DIST', [stack[0], stack[1], stack[2], true]),
        'LOGNORM.INV': () => {
            const p = Number(gVal(stack[0]));
            const mu = Number(gVal(stack[1]));
            const sigma = Number(gVal(stack[2]));
            if (![p, mu, sigma].every(Number.isFinite)) throw new Error('#VALUE!');
            if (sigma <= 0) throw new Error('#NUM!');
            return Math.exp(mu + sigma * normInv(p));
        },
        LOGINV: () => execute('LOGNORM.INV', stack),
        GAUSS: () => {
            const z = Number(gVal(stack[0]));
            if (!Number.isFinite(z)) throw new Error('#VALUE!');
            return normCdf(z) - 0.5;
        },
        'CONFIDENCE.NORM': () => {
            const alpha = Number(gVal(stack[0]));
            const sigma = Number(gVal(stack[1]));
            const size = Number(gVal(stack[2]));
            if (![alpha, sigma, size].every(Number.isFinite) || alpha <= 0 || alpha >= 1 || sigma <= 0 || size < 1) throw new Error('#NUM!');
            return normInv(1 - alpha / 2) * sigma / Math.sqrt(size);
        },
        CONFIDENCE: () => execute('CONFIDENCE.NORM', stack),
        'CONFIDENCE.T': () => execute('CONFIDENCE.NORM', stack),
        'EXPON.DIST': () => {
            const x = Number(gVal(stack[0]));
            const lambda = Number(gVal(stack[1]));
            const cumulative = !!gVal(stack[2]);
            if (![x, lambda].every(Number.isFinite)) throw new Error('#VALUE!');
            if (x < 0 || lambda <= 0) throw new Error('#NUM!');
            return cumulative ? 1 - Math.exp(-lambda * x) : lambda * Math.exp(-lambda * x);
        },
        EXPONDIST: () => execute('EXPON.DIST', stack),
        'POISSON.DIST': () => {
            const x = ensureInteger(gVal(stack[0]), '#NUM!');
            const meanV = Number(gVal(stack[1]));
            const cumulative = !!gVal(stack[2]);
            if (!Number.isFinite(meanV) || meanV <= 0 || x < 0) throw new Error('#NUM!');
            const pmf = (k) => Math.exp(-meanV + k * Math.log(meanV) - lnGamma(k + 1));
            if (!cumulative) return pmf(x);
            let total = 0;
            for (let k = 0; k <= x; k++) total += pmf(k);
            return total;
        },
        POISSON: () => execute('POISSON.DIST', stack),


        'BINOM.DIST': () => {
            const numberS = gVal(stack[0]);
            const trials = gVal(stack[1]);
            const prob = gVal(stack[2]);
            const cumulative = !!gVal(stack[3]);
            const k = ensureInteger(numberS, '#NUM!');
            if (!cumulative) return binomPmf(k, trials, prob);
            const n = ensureInteger(trials, '#NUM!');
            let total = 0;
            for (let i = 0; i <= k; i++) total += binomPmf(i, n, prob);
            return total;
        },
        BINOMDIST: () => execute('BINOM.DIST', stack),
        'BINOM.DIST.RANGE': () => {
            const trials = ensureInteger(gVal(stack[0]), '#NUM!');
            const pVal = Number(gVal(stack[1]));
            const numberS = ensureInteger(gVal(stack[2]), '#NUM!');
            const numberS2 = stack[3] !== undefined ? ensureInteger(gVal(stack[3]), '#NUM!') : numberS;
            if (!Number.isFinite(pVal) || pVal < 0 || pVal > 1 || trials < 0 || numberS < 0 || numberS2 < numberS || numberS2 > trials) throw new Error('#NUM!');
            let total = 0;
            for (let i = numberS; i <= numberS2; i++) total += binomPmf(i, trials, pVal);
            return total;
        },
        'BINOM.INV': () => {
            const trials = ensureInteger(gVal(stack[0]), '#NUM!');
            const pVal = Number(gVal(stack[1]));
            const alpha = Number(gVal(stack[2]));
            if (!Number.isFinite(pVal) || pVal < 0 || pVal > 1 || !Number.isFinite(alpha) || alpha < 0 || alpha > 1 || trials < 0) throw new Error('#NUM!');
            let cumulative = 0;
            for (let i = 0; i <= trials; i++) {
                cumulative += binomPmf(i, trials, pVal);
                if (cumulative >= alpha) return i;
            }
            return trials;
        },
        CRITBINOM: () => execute('BINOM.INV', stack),
        'NEGBINOM.DIST': () => {
            const failures = ensureInteger(gVal(stack[0]), '#NUM!');
            const successes = ensureInteger(gVal(stack[1]), '#NUM!');
            const pVal = Number(gVal(stack[2]));
            const cumulative = !!gVal(stack[3]);
            if (!cumulative) return negBinomPmf(failures, successes, pVal);
            let total = 0;
            for (let i = 0; i <= failures; i++) total += negBinomPmf(i, successes, pVal);
            return total;
        },
        NEGBINOMDIST: () => execute('NEGBINOM.DIST', [stack[0], stack[1], stack[2], false]),
        'HYPGEOM.DIST': () => {
            const sampleS = ensureInteger(gVal(stack[0]), '#NUM!');
            const numberSample = ensureInteger(gVal(stack[1]), '#NUM!');
            const populationS = ensureInteger(gVal(stack[2]), '#NUM!');
            const numberPop = ensureInteger(gVal(stack[3]), '#NUM!');
            const cumulative = !!gVal(stack[4]);
            if (numberPop <= 0 || numberSample < 0 || populationS < 0 || populationS > numberPop || sampleS < 0 || sampleS > numberSample) throw new Error('#NUM!');
            const minK = Math.max(0, numberSample - (numberPop - populationS));
            const maxK = Math.min(numberSample, populationS);
            const pmf = (k) => {
                if (k < minK || k > maxK) return 0;
                return (combination(populationS, k) * combination(numberPop - populationS, numberSample - k)) / combination(numberPop, numberSample);
            };
            if (!cumulative) return pmf(sampleS);
            let total = 0;
            for (let k = minK; k <= sampleS; k++) total += pmf(k);
            return total;
        },
        HYPGEOMDIST: () => execute('HYPGEOM.DIST', [stack[0], stack[1], stack[2], stack[3], false]),
        'WEIBULL.DIST': () => {
            const x = Number(gVal(stack[0]));
            const alpha = Number(gVal(stack[1]));
            const beta = Number(gVal(stack[2]));
            const cumulative = !!gVal(stack[3]);
            if (![x, alpha, beta].every(Number.isFinite)) throw new Error('#VALUE!');
            if (x < 0 || alpha <= 0 || beta <= 0) throw new Error('#NUM!');
            const term = (x / beta) ** alpha;
            if (cumulative) return 1 - Math.exp(-term);
            if (x === 0 && alpha < 1) throw new Error('#NUM!');
            return (alpha / beta) * ((x / beta) ** (alpha - 1)) * Math.exp(-term);
        },
        WEIBULL: () => execute('WEIBULL.DIST', stack),
        PROB: () => {
            const xRange = flattenAny(stack[0]);
            const pRange = flattenAny(stack[1]);
            const lower = Number(gVal(stack[2]));
            const upper = stack[3] !== undefined ? Number(gVal(stack[3])) : lower;
            if (xRange.length !== pRange.length || !Number.isFinite(lower) || !Number.isFinite(upper)) throw new Error('#N/A');
            let pSum = 0;
            let out = 0;
            for (let i = 0; i < xRange.length; i++) {
                const pVal = Number(pRange[i]);
                const x = Number(xRange[i]);
                if (!Number.isFinite(pVal) || pVal < 0 || pVal > 1) throw new Error('#NUM!');
                pSum += pVal;
                if (Number.isFinite(x) && x >= Math.min(lower, upper) && x <= Math.max(lower, upper)) out += pVal;
            }
            if (Math.abs(pSum - 1) > 1e-7) throw new Error('#NUM!');
            return out;
        },
        FREQUENCY: () => {
            const data = flattenNumeric(stack[0]);
            const bins = flattenNumeric(stack[1]).sort((a, b) => a - b);
            const counts = Array.from({ length: bins.length + 1 }, () => 0);
            data.forEach((value) => {
                let idx = bins.findIndex((b) => value <= b);
                if (idx < 0) idx = bins.length;
                counts[idx]++;
            });
            return counts.map((count) => [count]);
        },
        TRIMMEAN: () => {
            const values = gNums([stack[0]]).filter((v) => Number.isFinite(v)).sort((a, b) => a - b);
            const percent = Number(gVal(stack[1]));
            if (!Number.isFinite(percent) || percent < 0 || percent >= 1) throw new Error('#NUM!');
            if (!values.length) throw new Error('#DIV/0!');
            const trimEach = Math.floor(values.length * percent / 2);
            if (trimEach * 2 >= values.length) throw new Error('#NUM!');
            const trimmed = values.slice(trimEach, values.length - trimEach);
            return mean(trimmed);
        },


        SKEW: () => {
            const values = gNums(stack).filter((v) => Number.isFinite(v));
            const n = values.length;
            if (n < 3) throw new Error('#DIV/0!');
            const avg = mean(values);
            const s = Math.sqrt(values.reduce((sum, x) => sum + ((x - avg) ** 2), 0) / (n - 1));
            if (s === 0) throw new Error('#DIV/0!');
            const m3 = values.reduce((sum, x) => sum + ((x - avg) / s) ** 3, 0);
            return (n * m3) / ((n - 1) * (n - 2));
        },
        'SKEW.P': () => {
            const values = gNums(stack).filter((v) => Number.isFinite(v));
            const n = values.length;
            if (n < 1) throw new Error('#DIV/0!');
            const avg = mean(values);
            const sigma = Math.sqrt(values.reduce((sum, x) => sum + ((x - avg) ** 2), 0) / n);
            if (sigma === 0) throw new Error('#DIV/0!');
            return values.reduce((sum, x) => sum + ((x - avg) / sigma) ** 3, 0) / n;
        },
        KURT: () => {
            const values = gNums(stack).filter((v) => Number.isFinite(v));
            const n = values.length;
            if (n < 4) throw new Error('#DIV/0!');
            const avg = mean(values);
            const s2 = values.reduce((sum, x) => sum + ((x - avg) ** 2), 0) / (n - 1);
            if (s2 === 0) throw new Error('#DIV/0!');
            const s4 = s2 * s2;
            const sum4 = values.reduce((sum, x) => sum + ((x - avg) ** 4), 0);
            const term1 = (n * (n + 1) * sum4) / ((n - 1) * (n - 2) * (n - 3) * s4);
            const term2 = (3 * ((n - 1) ** 2)) / ((n - 2) * (n - 3));
            return term1 - term2;
        },
        STEYX: () => {
            const pairs = collectNumericPairs(stack[1], stack[0]);
            if (pairs.length < 3) throw new Error('#DIV/0!');
            const { slope, intercept } = calcSlopeIntercept(pairs);
            const sse = pairs.reduce((sum, [x, y]) => {
                const fit = intercept + slope * x;
                return sum + ((y - fit) ** 2);
            }, 0);
            return Math.sqrt(sse / (pairs.length - 2));
        },
        'FORECAST.LINEAR': () => {
            const x = Number(gVal(stack[0]));
            const pairs = collectNumericPairs(stack[2], stack[1]);
            const { slope, intercept } = calcSlopeIntercept(pairs);
            return intercept + slope * x;
        },
        FORECAST: () => execute('FORECAST.LINEAR', stack),
        LINEST: () => {
            const knownY = flattenNumeric(stack[0]);
            const knownX = toSeries(stack[1], knownY.length);
            if (knownY.length !== knownX.length || knownY.length < 2) throw new Error('#N/A');
            const pairs = knownX.map((x, i) => [x, knownY[i]]).filter(([x, y]) => Number.isFinite(x) && Number.isFinite(y));
            const keepConst = stack[2] !== undefined ? !!gVal(stack[2]) : true;
            const model = regressionFromPairs(pairs, !keepConst);
            return [[model.slope, model.intercept]];
        },
        LOGEST: () => {
            const knownY = flattenNumeric(stack[0]);
            if (knownY.some((y) => y <= 0)) throw new Error('#NUM!');
            const knownX = toSeries(stack[1], knownY.length);
            if (knownY.length !== knownX.length || knownY.length < 2) throw new Error('#N/A');
            const pairs = knownX.map((x, i) => [x, Math.log(knownY[i])]).filter(([x, y]) => Number.isFinite(x) && Number.isFinite(y));
            const keepConst = stack[2] !== undefined ? !!gVal(stack[2]) : true;
            const model = regressionFromPairs(pairs, !keepConst);
            return [[Math.exp(model.slope), Math.exp(model.intercept)]];
        },
        TREND: () => {
            const knownY = flattenNumeric(stack[0]);
            const knownX = toSeries(stack[1], knownY.length);
            if (knownY.length !== knownX.length || knownY.length < 2) throw new Error('#N/A');
            const pairs = knownX.map((x, i) => [x, knownY[i]]).filter(([x, y]) => Number.isFinite(x) && Number.isFinite(y));
            const keepConst = stack[3] !== undefined ? !!gVal(stack[3]) : true;
            const model = regressionFromPairs(pairs, !keepConst);
            const newX = toSeries(stack[2], knownY.length);
            const predicted = newX.map((x) => model.intercept + model.slope * x);
            if (stack[2] === undefined) return predicted.map((v) => [v]);
            return reshapeLikeArg(stack[2], predicted);
        },
        GROWTH: () => {
            const knownY = flattenNumeric(stack[0]);
            if (knownY.some((y) => y <= 0)) throw new Error('#NUM!');
            const knownX = toSeries(stack[1], knownY.length);
            if (knownY.length !== knownX.length || knownY.length < 2) throw new Error('#N/A');
            const pairs = knownX.map((x, i) => [x, Math.log(knownY[i])]).filter(([x, y]) => Number.isFinite(x) && Number.isFinite(y));
            const keepConst = stack[3] !== undefined ? !!gVal(stack[3]) : true;
            const model = regressionFromPairs(pairs, !keepConst);
            const newX = toSeries(stack[2], knownY.length);
            const predicted = newX.map((x) => Math.exp(model.intercept + model.slope * x));
            if (stack[2] === undefined) return predicted.map((v) => [v]);
            return reshapeLikeArg(stack[2], predicted);
        },
        'Z.TEST': () => {
            const values = flattenNumeric(stack[0]);
            const x = Number(gVal(stack[1]));
            const sigmaArg = stack[2] !== undefined ? Number(gVal(stack[2])) : null;
            if (!values.length || !Number.isFinite(x)) throw new Error('#VALUE!');
            const avg = mean(values);
            const sigma = Number.isFinite(sigmaArg)
                ? sigmaArg
                : Math.sqrt(values.reduce((sum, n) => sum + ((n - avg) ** 2), 0) / Math.max(values.length - 1, 1));
            if (!Number.isFinite(sigma) || sigma <= 0) throw new Error('#DIV/0!');
            const z = (avg - x) / (sigma / Math.sqrt(values.length));
            return 1 - normCdf(z);
        },
        ZTEST: () => execute('Z.TEST', stack),


        IMCOSH: () => {
            const z = gVal(stack[0]);
            const negZ = execute('IMPRODUCT', [-1, z]);
            const ez = execute('IMEXP', [z]);
            const enz = execute('IMEXP', [negZ]);
            return execute('IMDIV', [execute('IMSUM', [ez, enz]), 2]);
        },
        IMSINH: () => {
            const z = gVal(stack[0]);
            const negZ = execute('IMPRODUCT', [-1, z]);
            const ez = execute('IMEXP', [z]);
            const enz = execute('IMEXP', [negZ]);
            return execute('IMDIV', [execute('IMSUB', [ez, enz]), 2]);
        },
        IMTAN: () => execute('IMDIV', [execute('IMSIN', [gVal(stack[0])]), execute('IMCOS', [gVal(stack[0])])]),
        IMCOT: () => execute('IMDIV', [execute('IMCOS', [gVal(stack[0])]), execute('IMSIN', [gVal(stack[0])])]),
        IMSEC: () => execute('IMDIV', [1, execute('IMCOS', [gVal(stack[0])])]),
        IMCSC: () => execute('IMDIV', [1, execute('IMSIN', [gVal(stack[0])])]),
        IMSECH: () => execute('IMDIV', [1, execute('IMCOSH', [gVal(stack[0])])]),
        IMCSCH: () => execute('IMDIV', [1, execute('IMSINH', [gVal(stack[0])])]),

        STDEVP: () => execute('STDEV.P', stack),
        VARP: () => execute('VAR.P', stack),
    };
}

export default createMathStatFns;
