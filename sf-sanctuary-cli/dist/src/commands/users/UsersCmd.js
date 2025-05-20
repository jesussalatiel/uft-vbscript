"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.UsersCmd = exports.UserRole = void 0;
const child_process_1 = require("child_process");
const Logger_1 = require("../Logger");
var UserRole;
(function (UserRole) {
    UserRole["STANDARD_USER"] = "Standard User";
    UserRole["SYSTEM_ADMINISTRATOR"] = "System Administrator";
})(UserRole || (exports.UserRole = UserRole = {}));
class UsersCmd {
    async create(user) {
        const alias = user.username.slice(0, 5).toLowerCase();
        const profileData = await this.findByProfileRole(this.roleArgs(user.role));
        const profileId = profileData?.result?.records?.[0]?.Id;
        if (!profileId) {
            const alert = `Profile ID not found for role: ${this.roleArgs(user.role)}`;
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
        Logger_1.logger.success(JSON.stringify(result, null, 2));
    }
    async findByEmail(user) {
        const query = `SELECT Id, Name, Profile.Name FROM User WHERE Username = '${user.email}'`;
        const result = await this.runQuery(query);
        Logger_1.logger.success(JSON.stringify(result, null, 2));
    }
    async delete(user) {
        const currentUser = await this.getCurrentUser();
        const command = `sf data delete record -s User -i ${user.id} --target-org commondev`;
        const result = await this.runCommand(command);
        const parse = JSON.parse(result);
        if (parse.message.includes("INSUFFICIENT_ACCESS_OR_READONLY")) {
            const alert = `User '${currentUser}' does not have permission to delete users. Contact your admin.`;
            throw new Error(alert);
        }
        Logger_1.logger.success(`User with ID ${user.id} deleted successfully.`);
    }
    async getCurrentUser() {
        const command = `sf org display --target-org commondev --json`;
        const result = await this.runCommand(command);
        const parse = JSON.parse(result);
        return parse?.result?.username;
    }
    async findByProfileRole(role) {
        const query = `SELECT Id, Name FROM Profile WHERE Name = '${role}'`;
        const response = await this.runQuery(query);
        return JSON.parse(response);
    }
    async runCommand(command) {
        try {
            return (0, child_process_1.execSync)(command, { encoding: "utf-8" });
        }
        catch (error) {
            throw new Error(error.message);
        }
    }
    async runQuery(query) {
        const command = `sf data query --query "${query}" --target-org commondev --json`;
        return await this.runCommand(command);
    }
    roleArgs(role) {
        switch (role) {
            case "admin":
                return UserRole.SYSTEM_ADMINISTRATOR;
                break;
            default:
                return UserRole.STANDARD_USER;
        }
    }
}
exports.UsersCmd = UsersCmd;
