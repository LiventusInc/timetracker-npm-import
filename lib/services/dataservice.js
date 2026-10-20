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
const XLSX = __importStar(require("xlsx"));
const apiservice_1 = require("../services/apiservice");
const config = require('../../config.json');
let DataService = class DataService {
    constructor(apiService) {
        this.apiService = apiService;
    }
    getMappingData(file, columnsFromExcel, apiUrl, defaultOrganization) {
        var result = new Map();
        var errors = new Array();
        var wb = XLSX.readFile(file);
        var sheet_name_list = wb.SheetNames;
        var opt = columnsFromExcel ? { raw: false } : { header: "A", raw: false };
        var xlData = XLSX.utils.sheet_to_json(wb.Sheets[sheet_name_list[0]], opt);
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
                    workLog.timeStamp = new Date(date).toLocaleString();
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
                        workLog.timeStamp = new Date(date).toLocaleString();
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
