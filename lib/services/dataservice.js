"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};
var __param = (this && this.__param) || function (paramIndex, decorator) {
    return function (target, key) { decorator(target, key, paramIndex); }
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.DataService = void 0;
const inversify_1 = require("inversify");
const ExcelJS = __importStar(require("exceljs"));
const apiservice_1 = require("../services/apiservice");
const config = require('../../config.json');
// Reads the first sheet into one object per non-blank row, keyed by column letter, or by the
// header row's names when columnsFromExcel is set. Every value is display text, because the
// mapping below parses strings (dates via Date.parse, comma decimals via replace).
async function readRows(file, columnsFromExcel) {
    var wb = new ExcelJS.Workbook();
    var sheet;
    if (file.toLowerCase().endsWith('.csv')) {
        // Keep CSV fields as written; ExcelJS would otherwise convert numbers and dates.
        sheet = await wb.csv.readFile(file, { map: (value) => value });
    }
    else {
        await wb.xlsx.readFile(file);
        sheet = wb.worksheets[0];
    }
    var rows = [];
    var headers;
    sheet?.eachRow((row) => {
        var values = new Map();
        row.eachCell((cell, col) => {
            var text = cellText(cell.value);
            if (text !== undefined)
                values.set(col, text);
        });
        if (values.size === 0)
            return;
        if (columnsFromExcel && !headers) {
            headers = values;
            return;
        }
        var element = {};
        values.forEach((text, col) => {
            var key = columnsFromExcel ? headers.get(col) : columnLetter(col);
            if (key !== undefined)
                element[key] = text;
        });
        rows.push(element);
    });
    return rows;
}
function cellText(value) {
    if (value === null || value === undefined)
        return undefined;
    if (value instanceof Date)
        return dateText(value);
    if (typeof value === 'object') {
        if ('richText' in value)
            return cellText(value.richText.map(r => r.text).join(''));
        if ('formula' in value || 'sharedFormula' in value)
            return cellText(value.result);
        if ('text' in value)
            return cellText(value.text);
        if ('error' in value)
            return value.error;
        return undefined;
    }
    if (typeof value === 'boolean')
        return value ? 'TRUE' : 'FALSE';
    var text = String(value);
    return text === '' ? undefined : text;
}
// ExcelJS returns a date cell as a Date whose UTC fields are the wall-clock time Excel shows.
// Render it as text Date.parse reads back as local time: date only, time only (Excel's
// 1899-12-30 epoch day), or both.
function dateText(d) {
    var date = `${d.getUTCMonth() + 1}/${d.getUTCDate()}/${d.getUTCFullYear()}`;
    var hours = d.getUTCHours();
    var time = `${hours % 12 || 12}:${String(d.getUTCMinutes()).padStart(2, '0')}:${String(d.getUTCSeconds()).padStart(2, '0')} ${hours < 12 ? 'AM' : 'PM'}`;
    var timeOnly = d.getUTCFullYear() === 1899 && d.getUTCMonth() === 11 && d.getUTCDate() === 30;
    if (timeOnly)
        return time;
    if (hours === 0 && d.getUTCMinutes() === 0 && d.getUTCSeconds() === 0)
        return date;
    return `${date} ${time}`;
}
// The API's timeStamp is an ISO 8601 date-time read as local time in a 7pace user's time zone,
// so send the wall-clock time without an offset.
function apiTimestamp(d) {
    var pad = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}
function columnLetter(col) {
    var letters = '';
    for (; col > 0; col = Math.floor((col - 1) / 26)) {
        letters = String.fromCharCode(65 + (col - 1) % 26) + letters;
    }
    return letters;
}
let DataService = class DataService {
    constructor(apiService) {
        this.apiService = apiService;
    }
    async getMappingData(file, columnsFromExcel, apiUrl, defaultOrganization) {
        var result = new Map();
        var errors = new Array();
        var xlData = await readRows(file, columnsFromExcel);
        var map = columnsFromExcel
            ? { userName: 'UserName', workItem: 'WorkItem', timeStamp: 'TimeStamp', duration: 'Duration',
                billableDuration: 'BillableDuration', comment: 'Comment', activityType: 'ActivityType',
                date: 'Date', start: 'Start', organization: 'Organization' }
            : config.map;
        var addIndex = columnsFromExcel ? 2 : 1;
        let requireOrganization = !apiUrl;
        for (let index = 0; index < xlData.length; index++) {
            const element = xlData[index];
            var workLog = {};
            if (map.userName && element[map.userName]) {
                workLog.userId = element[map.userName];
            }
            if (map.workItem && element[map.workItem]) {
                var number = Number(element[map.workItem]);
                if (!isNaN(number) && Number.isInteger(number) && number > 0)
                    workLog.workItemId = number;
                else
                    errors.push(`"WorkItem" is not a valid value in a cell ${map.workItem}${index + addIndex}. The value may be not specified or must be an integer and greater than 0.`);
            }
            if (map.timeStamp && element[map.timeStamp]) {
                var date = Date.parse(element[map.timeStamp]);
                if (!isNaN(date))
                    workLog.timeStamp = apiTimestamp(new Date(date));
                else
                    errors.push(`"TimeStamp" is not a valid datetime value in a cell ${map.timeStamp}${index + addIndex}.`);
            }
            else {
                if (map.date && element[map.date]) {
                    var timestamp = element[map.date];
                    if (map.start && element[map.start]) {
                        timestamp += " " + element[map.start];
                    }
                    else {
                        timestamp += " 12:00 AM";
                    }
                    var date = Date.parse(timestamp);
                    if (!isNaN(date)) {
                        workLog.timeStamp = apiTimestamp(new Date(date));
                    }
                    else {
                        errors.push(`"Date" and "Start" do not specify a valid starting datetime value "${timestamp}" in a cell(s) ${map.date}${index + addIndex}.`);
                    }
                }
            }
            if (map.duration && element[map.duration]) {
                var number = Number(element[map.duration].replace(/,/g, '.'));
                if (!isNaN(number) && number > 0)
                    workLog.length = Math.ceil(number * 3600);
                else
                    errors.push(`"Duration" is not a valid value in a cell ${map.duration}${index + addIndex}. The value may be not specified or must be greater than 0.`);
            }
            if (map.billableDuration && element[map.billableDuration]) {
                var number = Number(element[map.billableDuration].replace(/,/g, '.'));
                if (!isNaN(number) && number >= 0)
                    workLog.billableLength = Math.ceil(number * 3600);
                else
                    errors.push(`"BillableDuration" is not a valid value in a cell ${map.billableDuration}${index + addIndex}. The value may be not specified or must be greater or equal than 0.`);
            }
            if (map.activityType && element[map.activityType]) {
                workLog.activityTypeId = element[map.activityType];
            }
            if (map.comment && element[map.comment]) {
                workLog.comment = element[map.comment];
            }
            if ((!workLog.comment || workLog.comment.trim() === '') && !workLog.workItemId) {
                errors.push(`"Comment" (cell ${map.comment}${index + addIndex}) and "WorkItem" (cell ${map.workItem}${index + addIndex}) are not specified. At least one of these values must be specified.`);
            }
            let mapKey;
            if (map.organization && element[map.organization] && element[map.organization] !== '') {
                mapKey = element[map.organization];
            }
            else if (requireOrganization) {
                if (!defaultOrganization || defaultOrganization.trim() === '') {
                    errors.push(`"Organization" (cell ${map.comment}${index + addIndex}) and -o / --organization <organization> command line parameter are not specified. At least one of these values must be specified.`);
                    continue;
                }
                else {
                    mapKey = defaultOrganization;
                }
            }
            else {
                mapKey = apiUrl;
            }
            if (mapKey) {
                mapKey = mapKey.toLowerCase();
            }
            let orgResult = result.get(mapKey);
            let organizationWorkLogs = orgResult ? orgResult : new Array();
            if (!result.has(mapKey)) {
                result.set(mapKey, organizationWorkLogs);
            }
            organizationWorkLogs.push(workLog);
        }
        if (errors.length > 0) {
            throw errors;
        }
        return result;
    }
    async getUserIdActivityTypeId(worklogs, apiUrl, auth, token, organization, user, password) {
        var users = new Map();
        var activityTypes = new Map();
        if (worklogs.some((w) => w.userId)) {
            users = await this.apiService.getUsers(apiUrl, auth, token, organization, user, password);
        }
        if (worklogs.some((w) => w.activityTypeId)) {
            activityTypes = await this.apiService.getActivityTypes(apiUrl, auth, token, organization, user, password);
        }
        var errors = new Set();
        for (let index = 0; index < worklogs.length; index++) {
            var worklog = worklogs[index];
            if (worklog.userId) {
                if (users.has(worklog.userId)) {
                    worklog.userId = users.get(worklog.userId);
                }
                else {
                    var err = `User "${worklog.userId}" does not exist in your organization`;
                    if (!errors.has(err))
                        errors.add(err);
                }
            }
            if (worklog.activityTypeId) {
                if (activityTypes.has(worklog.activityTypeId)) {
                    worklog.activityTypeId = activityTypes.get(worklog.activityTypeId);
                }
                else {
                    var err = `Activity type "${worklog.activityTypeId}" does not exist in your organization`;
                    if (!errors.has(err))
                        errors.add(err);
                }
            }
        }
        if (errors.size > 0) {
            throw Array.from(errors);
        }
        return worklogs;
    }
    async sendWorklogs(worklogs, apiUrl, auth, token, organization, user, password) {
        await this.apiService.postWorklogsBatch(worklogs, true, apiUrl, auth, token, organization, user, password);
        await this.apiService.postWorklogsBatch(worklogs, false, apiUrl, auth, token, organization, user, password);
    }
};
exports.DataService = DataService;
exports.DataService = DataService = __decorate([
    (0, inversify_1.injectable)(),
    __param(0, (0, inversify_1.inject)('APIService')),
    __metadata("design:paramtypes", [apiservice_1.APIService])
], DataService);
