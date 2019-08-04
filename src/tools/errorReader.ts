import { periods } from './utils'

export class ErrorReader {
    private sbData: DerbyJson.IGame
    private sbTemplate: IStatsbookTemplate
    private sbErrors: IErrorSummary
    private warningData: IWarningData

    constructor(sbData: DerbyJson.IGame,
                sbTemplate: IStatsbookTemplate,
                sbErrors: IErrorSummary,
                warningData: IWarningData) {
        this.sbData = sbData
        this.sbTemplate = sbTemplate
        this.sbErrors = sbErrors
        this.warningData = warningData
    }

    public processErrors(): void {
        periods.forEach((period) => {
            const jams = this.sbData.periods[period].jams
            const maxIdx = jams.length

            jams.forEach((jam, idx) => {
                const events = jam.events

                const penalties = events.filter((e) => e.event === 'penalty')
                const leadJammer = events.find((e) => e.event === 'lead')

                const thisJamEntries = events.filter((e) => e.event === 'enter box')
                let nextJam = null

                if (idx === maxIdx && period === '1') {
                    nextJam = this.sbData.periods['2'].jams[0]
                } else {
                    nextJam = jams[idx + 1]
                }
                const nextJamEntries = nextJam ?
                                nextJam.events.filter((e: DerbyJson.IEvent) => e.event === 'enter box')
                                : []

                penalties.forEach((penalty) => {
                    const inThis = thisJamEntries.find((e) => e.skater === penalty.skater)
                    const inNext = nextJamEntries.find((e) => e.skater === penalty.skater)

                    if (!inThis && !inNext) {
                        const skaterNum = penalty.skater.slice(5)

                        if (idx === maxIdx && period === '2') {
                            this.sbErrors.penalties.penaltyNoEntry.events
                            .push(`Team: ${penalty.team}, Period: ${period}, Jam: ${jam.number}, Skater: ${skaterNum}}`)
                        } else {
                            this.sbErrors.warnings.lastJamNoEntry.events
                            .push(`Team: ${penalty.team}, Period: ${period}, Jam: ${jam.number}, Skater: ${skaterNum}}`)
                        }

                        this.warningData.noEntries.push({
                            skater: penalty.skater,
                            team: penalty.team,
                            period,
                            jam: jam.number,
                        })
                    }
                })

                if (leadJammer) {
                    const jammerPenalty = penalties.find((e) => e.skater === leadJammer.skater)
                    const lostLead = events.find((e) => e.event === 'lost lead' && e.skater === leadJammer.skater)

                    if (lostLead && !jammerPenalty) {
                        const skaterNum = leadJammer.skater.slice(5)

                        this.sbErrors.warnings.leadPenaltyNotLost.events
                        .push(`Team: ${leadJammer.team}, Period: ${period}, Jam: ${jam.number}, Jammer: ${skaterNum}`)
                    }
                }
            })
        })
    }

    public processWarnings(): void {
    // Warning check: Possible substitution.
    // For each skater who has a $ or S without a corresponding penalty,
    // check to see if a different skater on the same team has
    // a penalty without a subsequent box exit.
        this.warningData.badStarts.forEach((bs) => {
            this.warningData.noEntries.find((ne) => ne.team === bs.team)
        })
    }
}
