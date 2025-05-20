"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.validateCustomer = validateCustomer;
const chalk_1 = __importDefault(require("chalk"));
const child_process_1 = require("child_process");
function validateCustomer(email) {
    try {
        // SELECT Id, Name, Profile.Name FROM User WHERE Username = 'wiwev22814@jazipo.com'
        const query = `SELECT Id, Name, Profile.Name FROM User WHERE Username = '${email}'`;
        const command = `sf data query --query "${query}" --target-org commondev`;
        const result = (0, child_process_1.execSync)(command, { encoding: "utf-8" });
        console.log(chalk_1.default.green(result));
        process.exit(0); // success
    }
    catch (error) {
        console.error(chalk_1.default.red("❌ Error executing Salesforce query:\n" + error.message));
        process.exit(1); // failure
    }
}
