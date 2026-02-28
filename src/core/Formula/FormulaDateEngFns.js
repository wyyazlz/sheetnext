export function createDateEngFns(ctx) {
    const {
        stack,
        gVal,
        toDate,
        execute,
        isoWeekNum,
        ensureInteger
    } = ctx;

    return {
        'CEILING.PRECISE': () => execute('CEILING_PRECISE', stack),
        'ISO.CEILING': () => execute('ISO_CEILING', stack),
        'FLOOR.PRECISE': () => execute('FLOOR_PRECISE', stack),
        'WORKDAY.INTL': () => execute('WORKDAY_INTL', stack),
        DAYS: () => {
            const endDate = toDate(gVal(stack[0]));
            const startDate = toDate(gVal(stack[1]));
            return Math.floor((endDate - startDate) / (1000 * 60 * 60 * 24));
        },
        'NETWORKDAYS.INTL': () => execute('NETWORKDAYS_INTL', stack),
        ISOWEEKNUM: () => isoWeekNum(toDate(gVal(stack[0]))),
        BITAND: () => {
            const a = BigInt(ensureInteger(gVal(stack[0]), '#NUM!'));
            const b = BigInt(ensureInteger(gVal(stack[1]), '#NUM!'));
            if (a < 0n || b < 0n) throw new Error('#NUM!');
            return Number(a & b);
        },
        BITOR: () => {
            const a = BigInt(ensureInteger(gVal(stack[0]), '#NUM!'));
            const b = BigInt(ensureInteger(gVal(stack[1]), '#NUM!'));
            if (a < 0n || b < 0n) throw new Error('#NUM!');
            return Number(a | b);
        },
        BITXOR: () => {
            const a = BigInt(ensureInteger(gVal(stack[0]), '#NUM!'));
            const b = BigInt(ensureInteger(gVal(stack[1]), '#NUM!'));
            if (a < 0n || b < 0n) throw new Error('#NUM!');
            return Number(a ^ b);
        },
        BITLSHIFT: () => {
            const n = BigInt(ensureInteger(gVal(stack[0]), '#NUM!'));
            const shift = ensureInteger(gVal(stack[1]), '#NUM!');
            if (n < 0n) throw new Error('#NUM!');
            return Number(shift >= 0 ? (n << BigInt(shift)) : (n >> BigInt(-shift)));
        },
        BITRSHIFT: () => {
            const n = BigInt(ensureInteger(gVal(stack[0]), '#NUM!'));
            const shift = ensureInteger(gVal(stack[1]), '#NUM!');
            if (n < 0n) throw new Error('#NUM!');
            return Number(shift >= 0 ? (n >> BigInt(shift)) : (n << BigInt(-shift)));
        },
        'ERF.PRECISE': () => execute('ERF', stack),
        'ERFC.PRECISE': () => execute('ERFC_PRECISE', stack),
    };
}

export default createDateEngFns;
