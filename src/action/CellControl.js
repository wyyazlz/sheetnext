export function insertCheckbox() {
    const sheet = this.SN.activeSheet;
    const areas = sheet.activeAreas?.length
        ? sheet.activeAreas
        : [{ s: sheet.activeCell, e: sheet.activeCell }];
    sheet.setCellControl(areas, { type: 'checkbox' });
    this.SN._r();
}

export function clearCheckbox() {
    const sheet = this.SN.activeSheet;
    const areas = sheet.activeAreas?.length
        ? sheet.activeAreas
        : [{ s: sheet.activeCell, e: sheet.activeCell }];
    sheet.clearCellControl(areas);
    this.SN._r();
}
