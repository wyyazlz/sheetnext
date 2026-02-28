export default class License {
    constructor() {
        this.isActivated = true;
        this.watermarkCanvas = null;
    }

    async activate() {
        this.isActivated = true;
        return true;
    }

    deactivate() {
        this.isActivated = true;
        return true;
    }

    getStatus() {
        return {
            isActivated: true,
            mode: 'open'
        };
    }

    _drawWatermark() {}
}
