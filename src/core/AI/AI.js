const OPEN_AI_MESSAGE = 'AI module is not available in SheetNext Open Edition.';

function notifyUnavailable(SN) {
    if (SN?.Utils?.toast) SN.Utils.toast(OPEN_AI_MESSAGE);
    return false;
}

export default class AI {
    constructor(SN) {
        this._SN = SN;
    }

    listenRequestStatus() {
        return () => {};
    }

    chatInput() {
        return notifyUnavailable(this._SN);
    }

    async conversation() {
        return notifyUnavailable(this._SN);
    }

    clearChat() {
        return notifyUnavailable(this._SN);
    }

    handleFileChange() {
        return notifyUnavailable(this._SN);
    }

    _handlePastedImage() {
        return false;
    }

    _previewImage() {
        return false;
    }
}
