"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.logger = void 0;
const chalk_1 = __importDefault(require("chalk"));
const console_1 = __importDefault(require("console"));
exports.logger = {
    json: (log) => {
        console_1.default.log(chalk_1.default.yellow(JSON.stringify(log, null, 2)));
    },
    success: (message) => {
        console_1.default.log(chalk_1.default.green(`✅ ${message}`));
    },
    error: (error) => {
        console_1.default.log(chalk_1.default.red(error));
    },
};
