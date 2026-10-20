#!/usr/bin/env node
"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.index = index;
require("reflect-metadata");
const inversify_1 = require("inversify");
const cli_1 = require("./cli");
const validatorservice_1 = require("./services/validatorservice");
const apiservice_1 = require("./services/apiservice");
const dataservice_1 = require("./services/dataservice");
function index() {
    const container = new inversify_1.Container();
    container.bind('ValidatorService').to(validatorservice_1.ValidatorService).inSingletonScope();
    container.bind('APIService').to(apiservice_1.APIService).inSingletonScope();
    container.bind('DataService').to(dataservice_1.DataService).inSingletonScope();
    container.bind('CLI').to(cli_1.CLI).inSingletonScope();
    return container.get('CLI');
}
index();
