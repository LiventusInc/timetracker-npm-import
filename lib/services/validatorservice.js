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
Object.defineProperty(exports, "__esModule", { value: true });
exports.ValidatorService = void 0;
const inversify_1 = require("inversify");
const fs = __importStar(require("fs"));
const authtype_1 = require("../models/authtype");
const colors = require("colors");
const config = require('../../config.json');
let ValidatorService = class ValidatorService {
    constructor() { }
    showSuccess(message) {
        console.log(colors.green(message));
    }
    showError(error) {
        console.error(colors.red(error));
    }
    showErrors(errors) {
        for (let index = 0; index < errors.length; index++) {
            console.error(colors.red(errors[index]));
        }
    }
    validate(program) {
        var errorsExist = false;
        if (!program.file) {
            this.showError('File is not specified');
            errorsExist = true;
        }
        else {
            try {
                if (!fs.existsSync(program.file)) {
                    this.showError(`File "${program.file}" does not exist.`);
                    errorsExist = true;
                }
                else {
                    var extFile = program.file.split('.').pop().toLowerCase();
                    if (!(extFile === "xlsx" || extFile === "csv")) {
                        this.showError(`File "${program.file}" is not in supported format(xlsx/csv).`);
                        errorsExist = true;
                    }
                }
            }
            catch (error) {
                this.showError(error);
                errorsExist = true;
            }
        }
        if (!(program.authorization.toLowerCase() in authtype_1.AuthType)) {
            this.showError('Unknown API authorization type. Supported types: "token", "ntlm"');
            errorsExist = true;
        }
        if (!config.map && !program.map) {
            this.showError("Map is not specified in config.json");
            errorsExist = true;
        }
        if (program.authorization.toLowerCase() == authtype_1.AuthType.ntlm.valueOf() && (!program.api || !program.organization)) {
            this.showError("Api url and organization are not specified");
            errorsExist = true;
        }
        if (program.authorization.toLowerCase() == authtype_1.AuthType.token.valueOf() && !program.token) {
            if (!config.organizationTokens || config.organizationTokens.length === 0) {
                this.showError("Token is not specified, and organizationTokens not specified in config.json");
                errorsExist = true;
            }
        }
        if (program.authorization.toLowerCase() == authtype_1.AuthType.ntlm.valueOf() && (!program.user || !program.password)) {
            this.showError("User and password are not specified");
            errorsExist = true;
        }
        return errorsExist;
    }
};
exports.ValidatorService = ValidatorService;
exports.ValidatorService = ValidatorService = __decorate([
    (0, inversify_1.injectable)(),
    __metadata("design:paramtypes", [])
], ValidatorService);
