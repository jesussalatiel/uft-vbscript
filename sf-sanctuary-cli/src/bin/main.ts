import { Command } from "commander";
import customers from "../commands/users";
import * as path from "path";
import * as fs from "fs";

const program = new Command();
const packageJsonPath = path.resolve(__dirname, "../../package.json");
const packageJson = JSON.parse(fs.readFileSync(packageJsonPath, "utf-8"));

program
  .name("sf-sanctuary-cli")
  .description("Salesforce Sanctuary CLI")
  .version(packageJson);

program.addCommand(customers);

program.parse(process.argv);
