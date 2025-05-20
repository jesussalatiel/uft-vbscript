"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const commander_1 = require("commander");
const UsersCmd_1 = require("./UsersCmd");
const customers = new commander_1.Command("customers");
customers
    .command("create")
    .description("Create Customer")
    .requiredOption("-u, --username <username>", "Username")
    .requiredOption("-e, --email <email>", "Email")
    .requiredOption("-l, --lastname <lastname>", "Last name")
    .requiredOption("-r, --role <role>", "Role")
    .action((options) => new UsersCmd_1.UsersCmd().create(options));
customers
    .command("delete")
    .description("Delete Salesforce user by ID")
    .requiredOption("-i, --id <userId>", "User ID to delete")
    .action((opts) => new UsersCmd_1.UsersCmd().delete(opts));
customers
    .command("find")
    .description("Find Customer")
    .requiredOption("--email <email>", "Customer email")
    .action((email) => new UsersCmd_1.UsersCmd().findByEmail(email));
exports.default = customers;
