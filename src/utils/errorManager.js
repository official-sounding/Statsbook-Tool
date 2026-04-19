const sbErrorTemplate = require('../assets/sberrors.json');



class ErrorManager {
    data = sbErrorTemplate;
    warningData = {
        badStarts: [],
        noEntries: [],
        badContinues: [],
        noExits: [],
        foulouts: [],
        expulsions: [],
        lost: [],
        jamsCalledInjury: [],
        lineupThree: []
    };

    addWarning(type, msg) {
        this.addEvent('warning', type, msg);

    }

    addScoreError(type, msg) {
        this.addEvent('score', type, msg);
    }

    addLineupError(type, msg) {
        this.addEvent('lineup', type, msg);
    }

    addPenaltyError(type, msg) {
        this.addEvent('penalty', type, msg);
    }

    addEvent(bucket, type, msg) {
        const outer = this.data[bucket];

        if(!outer) {
            throw new Error(`invalid bucket ${bucket}`);
        }

        const inner = outer[type];

        if(!inner) {
            throw new Error(`invalid ${bucket} type ${inner}`);
        }

        inner.events.push(msg);
    }

    reset() {
        Object.values(data).forEach((bucket) => {
            Object.values(bucket).forEach((type) => {
                type.events = [];
            });
        });

        Object.keys(this.warningData).forEach((bucket) => {
            this.warningData[bucket] = [];
        })
    }
}

module.exports = new ErrorManager();