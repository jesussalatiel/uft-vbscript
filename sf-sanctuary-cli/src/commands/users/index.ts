import { Command } from "commander";
import { UsersCmd } from "./UsersCmd";

const customers = new Command("customers");

customers
  .command("create")
  .description("Create Customer")
  .requiredOption("-u, --username <username>", "Username")
  .requiredOption("-e, --email <email>", "Email")
  .requiredOption("-l, --lastname <lastname>", "Last name")
  .requiredOption("-r, --role <role>", "Role")
  .action((options) => new UsersCmd().create(options));

customers
  .command("delete")
  .description("Delete Salesforce user by ID")
  .requiredOption("-i, --id <userId>", "User ID to delete")
  .action((opts) => new UsersCmd().delete(opts));

customers
  .command("find")
  .description("Find Customer")
  .requiredOption("--email <email>", "Customer email")
  .action((email) => new UsersCmd().findByEmail(email));

export default customers;
