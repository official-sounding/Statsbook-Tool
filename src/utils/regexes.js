const anSP = /^sp\*?$/i
const mySP = /^sp$/i
const anINJ = /^inj\*?$/i
const npRe = /(\d)\+NP/
const ippRe = /(\d)\+(\d)/
const jamNoRe = /^(\d+|SP|SP\*|INJ|INJ\*)$/i

module.exports = {
    anSP,
    mySP,
    anINJ,
    npRe,
    ippRe,
    jamNoRe
}