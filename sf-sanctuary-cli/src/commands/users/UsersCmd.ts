import { execSync } from "child_process";
import { logger } from "../Logger";

export enum UserRole {
  STANDARD_USER = "Standard User",
  SYSTEM_ADMINISTRATOR = "System Administrator",
}

interface User {
  id?: string;
  username: string;
  email: string;
  lastname: string;
  role: string;
}

export class UsersCmd {
  async create(user: User) {
    const alias = user.username.slice(0, 5).toLowerCase();

    const profileData = await this.findByProfileRole(this.roleArgs(user.role));

    const profileId = profileData?.result?.records?.[0]?.Id;

    if (!profileId) {
      const alert = `Profile ID not found for role: ${this.roleArgs(
        user.role
      )}`;

      throw new Error(alert);
    }

    const values = [
      `Username=${user.email}`,
      `Email=${user.email}`,
      `LastName=${user.lastname}`,
      `Alias=${alias}`,
      `TimeZoneSidKey=America/New_York`,
      `LocaleSidKey=en_US`,
      `EmailEncodingKey=UTF-8`,
      `LanguageLocaleKey=en_US`,
      `ProfileId=${profileId}`,
    ].join(" ");

    const command = `sf data create record --sobject User --values "${values}" --target-org commondev`;
    const result = await this.runCommand(command);

    logger.success(JSON.stringify(result, null, 2));
  }

  async findByEmail(user: User) {
    const query = `SELECT Id, Name, Profile.Name, IsActive FROM User WHERE Username = '${user.email}'`;
    const result = await this.runQuery(query, false);

    logger.success(result);
  }

  async delete(user: User) {
    // sf data update record --sobject User --record-id 005Pu00000FWKRCIA5 --values "IsActive=false" --target-org commondev
    const currentUser = await this.getCurrentUser();
    const command = `sf data delete record -s User -i ${user.id} --target-org commondev`;
    const result = await this.runCommand(command);
    const parse = JSON.parse(result!);

    if (parse.message.includes("INSUFFICIENT_ACCESS_OR_READONLY")) {
      const alert = `User '${currentUser}' does not have permission to delete users. Contact your admin.`;
      throw new Error(alert);
    }

    logger.success(`User with ID ${user.id} deleted successfully.`);
  }

  private async getCurrentUser() {
    const command = `sf org display --target-org commondev --json`;
    const result = await this.runCommand(command);
    const parse = JSON.parse(result!);
    return parse?.result?.username;
  }

  private async findByProfileRole(role: UserRole) {
    const query = `SELECT Id, Name FROM Profile WHERE Name = '${role}'`;
    const response = await this.runQuery(query);
    return JSON.parse(response!);
  }

  private async runCommand(command: string) {
    try {
      return execSync(command, { encoding: "utf-8" });
    } catch (error: any) {
      throw new Error(error.message);
    }
  }

  private async runQuery(query: string, isJson: boolean = true) {
    const format = isJson ? "--json" : "";
    const command = `sf data query --query "${query}" --target-org commondev ${format}`;
    return await this.runCommand(command);
  }

  private roleArgs(role: string) {
    switch (role) {
      case "admin":
        return UserRole.SYSTEM_ADMINISTRATOR;
      default:
        return UserRole.STANDARD_USER;
    }
  }
}
