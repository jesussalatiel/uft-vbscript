"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const commander_1 = require("commander");
const customers_1 = require("../commands/customers");
const program = new commander_1.Command();
program
    .name("mycli")
    .description("CLI for Salesforce & UFT tasks")
    .version("1.0.0");
program
    .command("validate-customer")
    .description("Verify that a customer has a contract")
    .requiredOption("--customer-id <id>", "Salesforce Customer ID")
    .action((options) => {
    (0, customers_1.validateCustomer)(options.customerId);
});
program.parse(process.argv);
