export function getOffset(x, y, w, h) {
    const rect = this.handleLayer.getBoundingClientRect();
    const zoom = this.activeSheet?.zoom ?? 1;
    const obj = { s: {}, e: {} };
    const { r: sr, c: sc } = this.getCellIndex({ clientX: rect.left + x * zoom, clientY: rect.top + y * zoom });
    const { r: er, c: ec } = this.getCellIndex({ clientX: rect.left + (x + w) * zoom, clientY: rect.top + (y + h) * zoom });
    obj.s.r = sr;
    obj.s.c = sc;
    obj.e.r = er;
    obj.e.c = ec;
    const sheet = this.activeSheet;
    const startInfo = sheet.getCellInViewInfo(sr, sc, false);
    const endInfo = sheet.getCellInViewInfo(er, ec, false);
    obj.s.offsetX = x - startInfo.x;
    obj.s.offsetY = y - startInfo.y;
    obj.e.offsetX = x + w - endInfo.x;
    obj.e.offsetY = y + h - endInfo.y;
    return obj;
}

export async function rDrawing() {}
