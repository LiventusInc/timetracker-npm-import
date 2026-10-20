"use strict";
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
var __metadata = (this && this.__metadata) || function (k, v) {
    if (typeof Reflect === "object" && typeof Reflect.metadata === "function") return Reflect.metadata(k, v);
};
var __param = (this && this.__param) || function (paramIndex, decorator) {
    return function (target, key) { decorator(target, key, paramIndex); }
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.CLI = void 0;
const inversify_1 = require("inversify");
const validatorservice_1 = require("./services/validatorservice");
const dataservice_1 = require("./services/dataservice");
const authtype_1 = require("./models/authtype");
const program = require('commander');
const config = require('../config.json');
let CLI = class CLI {
    constructor(validatorService, dataService) {
        this.validatorService = validatorService;
        this.dataService = dataService;
        this.executeCLI();
    }
    executeCLI() {
        program
            .name("import-worklogs")
            .version('0.0.2')
            .description("CLI for importing worklogs(csv/excel)")
            .option('-f, --file <file>', 'File path for importing(csv/excel)')
            .option('-a, --authorization [authorization]', 'API authorization type (token or ntlm). Default value is "token"', 'token')
            .option('-o, --organization <organization>', 'The name of your organization in Azure DevOps or collection in On-premise version')
            .option('--api [api]', 'The api url. This option is required for import to On-premise')
            .option('-t, --token [token]', 'Reporting API Access Token. This option is required for import to Azure DevOps')
            .option('-u, --user [user]', 'NTLM authorization username. This option is required for import import to On-premise version')
            .option('-p, --password [password]', 'NTLM authorization password. This option is required for import to On-premise version)')
            .option('-m, --map', 'Mapping worklog fields from column names(first row in Excel file)', false)
            .parse(process.argv);
        if (!process.argv.slice(2).length || this.validatorService.validate(program)) {
            program.outputHelp();
            process.exit(1);
        }
        //work with excel data and mapping columns
        var mapData = new Map();
        try {
            mapData = this.dataService.getMappingData(program.file, program.map, program.api, program.organization);
            if (mapData.size === 0) {
                throw 'No worklog entries were found in the file.';
            }
        }
        catch (error) {
            error instanceof Array ? this.validatorService.showErrors(error) : this.validatorService.showError(error);
            process.exit(1);
        }
        //api requests logic
        var authType = program.authorization.toLowerCase() == authtype_1.AuthType.token ? authtype_1.AuthType.token : authtype_1.AuthType.ntlm;
        mapData.forEach((worklogs, organization) => {
            const apiUrl = program.api ? program.api : `https://${organization}.timehub.7pace.com`;
            const token = program.token ? program.token : config.organizationTokens[organization];
            console.log(`Importing worklogs to ${apiUrl}...`);
            this.dataService.getUserIdActivityTypeId(worklogs, apiUrl, authType, token, organization, program.user, program.password)
                .then(() => {
                this.dataService.sendWorklogs(worklogs, apiUrl, authType, token, organization, program.user, program.password)
                    .then(() => this.validatorService.showSuccess("Worklogs import complete"))
                    .catch((error) => this.validatorService.showError(error));
            })
                .catch((error) => error instanceof Array ? this.validatorService.showErrors(error) : this.validatorService.showError(error));
        });
    }
};
exports.CLI = CLI;
exports.CLI = CLI = __decorate([
    (0, inversify_1.injectable)(),
    __param(0, (0, inversify_1.inject)('ValidatorService')),
    __param(1, (0, inversify_1.inject)('DataService')),
    __metadata("design:paramtypes", [validatorservice_1.ValidatorService,
        dataservice_1.DataService])
], CLI);
