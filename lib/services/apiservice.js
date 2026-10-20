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
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.APIService = void 0;
const inversify_1 = require("inversify");
const axios_1 = __importDefault(require("axios"));
const util = __importStar(require("util"));
const authtype_1 = require("../models/authtype");
const config = require('../../config.json');
var httpntlm = require('httpntlm');
const httpntlmGetAsync = util.promisify(httpntlm.get);
const httpntlmPostAsync = util.promisify(httpntlm.post);
let APIService = class APIService {
    constructor() {
        this.apiVersion = "3.0-preview";
    }
    async getUsers(apiUrl, auth, token, organization, user, password) {
        var users = new Map();
        try {
            var response;
            if (auth == authtype_1.AuthType.token) {
                var opt = { headers: { 'Authorization': `Bearer ${token}` }, params: { 'api-version': this.apiVersion } };
                response = (await axios_1.default.get(apiUrl + '/api/rest/users', opt)).data;
            }
            else {
                var res = await httpntlmGetAsync({
                    url: apiUrl + `/api/${organization}/rest/users?api-version=${this.apiVersion}`,
                    username: user,
                    password: password
                });
                response = JSON.parse(res.body);
                if (res.statusCode < 200 || res.statusCode > 299)
                    throw response.error;
            }
            if (response.data && Array.isArray(response.data)) {
                for (let index = 0; index < response.data.length; index++) {
                    const user = response.data[index];
                    if (user) {
                        var uniqueName = auth == authtype_1.AuthType.ntlm ? user.uniqueName.replace("\\", "\\\\") : user.uniqueName;
                        if (uniqueName && !users.has(uniqueName)) {
                            users.set(uniqueName, user.id);
                        }
                    }
                }
            }
        }
        catch (error) {
            if (error.response && error.response.data.error)
                throw error.response.data.error;
            else if (error.message)
                throw error.message;
            else
                throw error;
        }
        return users;
    }
    async getActivityTypes(apiUrl, auth, token, organization, user, password) {
        var activityTypes = new Map();
        try {
            var response;
            if (auth == authtype_1.AuthType.token) {
                var opt = { headers: { 'Authorization': `Bearer ${token}` }, params: { 'api-version': this.apiVersion } };
                response = (await axios_1.default.get(apiUrl + '/api/rest/activityTypes', opt)).data;
            }
            else {
                var res = await httpntlmGetAsync({
                    url: apiUrl + `/api/${organization}/rest/activityTypes?api-version=${this.apiVersion}`,
                    username: user,
                    password: password
                });
                response = JSON.parse(res.body);
                if (res.statusCode < 200 || res.statusCode > 299)
                    throw response.error;
            }
            if (response.data.activityTypes && Array.isArray(response.data.activityTypes)) {
                for (let index = 0; index < response.data.activityTypes.length; index++) {
                    const activityType = response.data.activityTypes[index];
                    if (activityType && activityType.name && !activityTypes.has(activityType.name)) {
                        activityTypes.set(activityType.name, activityType.id);
                    }
                }
            }
        }
        catch (error) {
            if (error.response && error.response.data.error)
                throw error.response.data.error;
            else if (error.message)
                throw error.message;
            else
                throw error;
        }
        return activityTypes;
    }
    async postWorklogsBatch(worklogs, validateOnly, apiUrl, auth, token, organization, user, password) {
        const batchesCount = Math.ceil(worklogs.length / 500);
        try {
            if (auth == authtype_1.AuthType.token) {
                var opt = { headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }, params: { 'api-version': this.apiVersion, 'validateOnly': validateOnly }, };
                var url = apiUrl + '/api/rest/workLogs/batch';
                for (let index = 0; index < batchesCount; index++) {
                    var batch = worklogs.slice(index * 500, (index + 1) * 500);
                    var res = await axios_1.default.post(url, batch, opt);
                }
            }
            else {
                for (let index = 0; index < batchesCount; index++) {
                    var batch = worklogs.slice(index * 500, (index + 1) * 500);
                    var resp = await httpntlmPostAsync({
                        url: apiUrl + `/api/${organization}/rest/workLogs/batch?api-version=${this.apiVersion}&validateOnly=${validateOnly}`,
                        username: user,
                        password: password,
                        body: JSON.stringify(batch),
                        headers: { 'Content-Type': 'application/json' }
                    });
                    var response = JSON.parse(resp.body);
                    if (resp.statusCode < 200 || resp.statusCode > 299)
                        throw response.error;
                }
            }
        }
        catch (error) {
            if (error.response && error.response.data.error)
                throw error.response.data.error;
            else if (error.message)
                throw error.message;
            else
                throw error;
        }
    }
};
exports.APIService = APIService;
exports.APIService = APIService = __decorate([
    (0, inversify_1.injectable)(),
    __metadata("design:paramtypes", [])
], APIService);
