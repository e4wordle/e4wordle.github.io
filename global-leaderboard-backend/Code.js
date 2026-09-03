let doGet = (e) => {
    e = e || { parameter: { dev: true, ignoreCache: true } };
    let lock;
    try {
        lock = getLock();
        let { dev, ignoreCache } = e.parameter;
        let sheet = getDataSheet(dev);
        if (ignoreCache) {
            CacheService.getScriptCache().remove(sheet.getName());
        }
        let ans = getLeaderboardData(sheet);
        console.log(JSON.stringify(ans));
        return jsonResponse(ans);
    } catch (error) {
        console.error(error.stack);
        return jsonResponse({ error: error.toString() });
    } finally {
        lock?.releaseLock();
    }
};

let doPost = (e) => {
    e = e || { postData: { contents: JSON.stringify(MOCK_POST_DATA) } };
    let lock;
    try {
        lock = getLock();
        let { dev, ...data } = JSON.parse(e.postData.contents);
        let sheet = getDataSheet(dev);
        sheet.appendRow([Date.now(), JSON.stringify(data)]);
        CacheService.getScriptCache().remove(sheet.getName());
        let ans = getLeaderboardData(sheet);
        console.log(JSON.stringify(ans));
        return jsonResponse(ans);
    } catch (error) {
        console.error(error.stack);
        return jsonResponse({ error: error.toString() });
    } finally {
        lock?.releaseLock();
    }
};

let removeOldRows = () => {
    let lock;
    try {
        lock = getLock();
        let sheet = getDataSheet(false);
        if (sheet.getLastRow() === 0) {
            return;
        }
        let rows = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
        let cutoff = Date.now() - 7 * 86400000;
        let i = 0;
        for (; i < rows.length; ++i) {
            if (rows[i][0] > cutoff) {
                break;
            }
        }
        if (i > 0) {
            sheet.deleteRows(1, i);
        }
    } catch (error) {
        console.error(error.stack);
    } finally {
        lock?.releaseLock();
    }
};

let jsonResponse = (x) => ContentService.createTextOutput(JSON.stringify(x)).setMimeType(ContentService.MimeType.JSON);

let compareArrays = (a, b) => {
    let n = Math.min(a.length, b.length);
    for (let i = 0; i < n; ++i) {
        if (a[i] < b[i]) {
            return -1;
        }
        if (a[i] > b[i]) {
            return 1;
        }
    }
    return a.length - b.length;
};

let getLock = () => {
    let lock = LockService.getScriptLock();
    lock.waitLock(12500);
    return lock;
};

let getDataSheet = (dev) => SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dev ? "dev_data" : "data");

let getLeaderboardData = (sheet) => {
    let cache = CacheService.getScriptCache();
    const VERSION = 1;
    try {
        let { version, ans } = JSON.parse(cache.get(sheet.getName()));
        if (version === VERSION) {
            return ans;
        }
    } catch {}
    let now = Date.now();
    let ans = { date: now, day: {}, week: {} };
    if (sheet.getLastRow() === 0) {
        return ans;
    }
    let rows = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();
    let soloRecords = new Map();
    let teamRecords = new Map();
    let modesSeen = new Map();
    for (let [date, data] of rows) {
        if (date < now - 7 * 86400000) {
            continue;
        }
        let { mode, opener, items } = JSON.parse(data);
        for (let item of items) {
            let { xp, failed, playtime, score } = item;
            let updateSoloRecord = (player, visible) => {
                let record = soloRecords.get(player);
                if (!record) {
                    record = {
                        day: {
                            overall: { score: [0, 0, date] }, // total XP, total challenges, last played
                        },
                        week: {
                            overall: { score: [0, 0, 0] }, // total XP, total challenges, total playtime
                        },
                    };
                    soloRecords.set(player, record);
                }
                record.visible = visible;
                record.week.overall.score[0] += xp;
                ++record.week.overall.score[1];
                record.week.overall.score[2] += playtime;
                if (!failed && compareArrays(score, record.week[mode]?.score || []) >= 0) {
                    record.week[mode] = { score: [...score, -date], opener };
                }
                if (date >= now - 86400000) {
                    record.day.overall.score[0] += xp;
                    ++record.day.overall.score[1];
                    record.day.overall.score[2] = date;
                    if (!failed && compareArrays(score, record.day[mode]?.score || []) >= 0) {
                        record.day[mode] = { score: [...score, -date], opener };
                    }
                }
            };
            let isSolo = "player" in item;
            modesSeen.set(mode, isSolo);
            if (isSolo) {
                updateSoloRecord(item.player, item.visible);
                continue;
            }
            let players = [];
            for (let { player, visible } of item.team) {
                updateSoloRecord(player, visible);
                players.push(player);
            }
            let teamKey = players.sort().join("\n");
            let record = teamRecords.get(teamKey);
            if (!record) {
                record = { day: {}, week: {} };
                teamRecords.set(teamKey, record);
            }
            if (!failed && compareArrays(score, record.week[mode]?.score || []) >= 0) {
                record.week[mode] = { score: [...score, -date], opener };
            }
            if (!failed && date >= now - 86400000 && compareArrays(score, record.day[mode]?.score || []) >= 0) {
                record.day[mode] = { score: [...score, -date], opener };
            }
        }
    }
    ans.day.players = 0;
    ans.week.players = 0;
    for (let record of soloRecords.values()) {
        if (record.visible) {
            ans.day.players += record.day.overall.score[0] > 0;
            ans.week.players += record.week.overall.score[0] > 0;
        }
    }
    modesSeen.set("overall", true);
    for (let period of ["day", "week"]) {
        for (let [mode, isSolo] of modesSeen) {
            let items = [];
            if (isSolo) {
                for (let [player, record] of soloRecords) {
                    if (!record.visible || !record[period][mode]?.score[0]) {
                        continue;
                    }
                    items.push({ player, ...record[period][mode] });
                }
            } else {
                for (let [teamKey, record] of teamRecords) {
                    let team = teamKey.split("\n");
                    if (!record[period][mode]?.score[0] || !team.some((p) => soloRecords.get(p).visible)) {
                        continue;
                    }
                    items.push({ team, ...record[period][mode] });
                }
            }
            items.sort((b, a) => compareArrays(a.score, b.score));
            ans[period][mode] = items.slice(0, 36);
        }
    }
    cache.put(sheet.getName(), JSON.stringify({ version: VERSION, ans }), 1800);
    return ans;
};

const MOCK_POST_DATA = [
    {
        dev: true,
        mode: "sprint",
        opener: "",
        items: [
            { player: "e4", visible: true, xp: 10, failed: false, playtime: 45, score: [-45, true] },
            { player: "d4", visible: true, xp: 9, failed: false, playtime: 54, score: [-54, true] },
        ],
    },
][0];
