import chalk from "chalk";
import Console, { error } from "console";

export const logger = {
  json: (log: any) => {
    Console.log(chalk.yellow(JSON.stringify(log, null, 2)));
  },

  success: (message: string) => {
    Console.log(chalk.green(message));
  },

  error: (error: any) => {
    Console.log(chalk.red(error));
  },
};
