
class PenaltyBox {
    home = new Set();
    away = new Set();

    enter(team, skater) {
        if(this[team].has(skater)) {
            return false;
        }

        this[team].add(skater);
        return true;
    }

    exit(team, skater) {
        if(!this[team].has(skater)) {
            return false;
        }

        this[team].delete(skater);
        return true;
    }

    has(team, skater) {
        return this[team].has(skater);
    }

    skaters(team) {
        return Array.from(this[team]);
    }
}

class DataManager {
    sbData = { periods: {'1': {jams: []}, '2': {jams: []}} };
    penalties = {}
    starPasses = {}

    addPenalty(skater, penalty) {
        if(!this.penalties[skater]) {
            this.penalties[skater] = [];
        }
        this.penalties[skater].push(penalty)
    }

    penaltyCount(skater) {
        return this.penalties[skater] ?? 0;
    }

    reset() {
        this.penalties = {};
        this.starPasses = {};
    }

    addStarPass(sp) {
        const key = `${sp.period}:${sp.jam}`;
        if(!this.starPasses[key]) {
            this.starPasses[key] = { 
                period: sp.period,
                jam: sp.jam,
                count: 0
            };
        }

        this.starPasses[key].count++;
    }

    unmatchedStarPasses() {
        return Object.values(this.starPasses).filter(sp => sp.count === 1);
    }
}

module.exports = { 
    instance: new DataManager(),
    PenaltyBox
};