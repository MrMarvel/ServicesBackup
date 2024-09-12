import json
import re
import subprocess
import sys

import pyuac


class Program:
    _only_save_rows = ["Name", "DisplayName", "StartType"]
    def __init__(self):
        self._codepage = str(self.current_codepage())

    @staticmethod
    def current_codepage() -> int:
        result = subprocess.getoutput("chcp")
        return int(re.search(r"\d+", result).group())

    def backup_services(self, filename):
        services = self.list_services()
        with open(filename, "w", encoding="utf-8") as f:
            json.dump(services, f, indent=4, ensure_ascii=False)

    def load_services_from_file(self, filename) -> list[dict]:
        with open(filename, "r", encoding="utf-8") as f:
            services = json.load(f)
        if not isinstance(services, list) or not all(isinstance(service, dict) for service in services):
            raise ValueError("File is not a list of services")
        for service in services:
            for key in self._only_save_rows:
                if key not in service:
                    raise ValueError(f"Service is missing key \"{key}\" in:\n"
                                     f"{json.dumps(service, indent=4, ensure_ascii=False)}")
        services_sorted = list(sorted(services, key=lambda x: x["DisplayName"]))
        return services_sorted

    def restore_services(self, filename):
        services = self.load_services_from_file(filename)
        for service in services:
            cmd = f"""
            Set-Service -Name "{service["Name"]}" -StartupType {service["StartType"]}
            """
            completed = subprocess.run(["pwsh", "-NoProfile", "-Command", cmd],
                                       capture_output=True, encoding=self._codepage, )
            if completed.returncode != 0:
                print(f"Failed to restore service {service['Name']}:")
                print(completed.stderr)
        print("Services restored")

    def print_services(self):
        services = self.list_services()
        print(json.dumps(services, indent=4, ensure_ascii=False))

    def list_services(self) -> list[dict]:
        cmd = """
        Get-Service | Select-Object -Property Name,DisplayName,Status,StartType | ConvertTo-Json -EnumsAsStrings
        """

        # result = subprocess.getoutput(f"pwsh -NoProfile -Command \"{cmd}\"")
        completed = subprocess.run(["pwsh", "-NoProfile", "-Command", cmd],
                                   capture_output=True, encoding=self._codepage, )
        if completed.returncode != 0:
            print("Failed to get some services:")
            print(completed.stderr)
        try:
            services_json_str = completed.stdout
            services: list[dict] = json.loads(services_json_str)
        except Exception as e:
            print("Error parsing services")
            print(completed.stdout)
            print(e)
            sys.exit(1)
        services_cut = [{k: v for k, v in service.items() if k in self._only_save_rows} for service in services]
        services_sorted = list(sorted(services_cut, key=lambda x: x["DisplayName"]))
        return services_sorted

    def configurations_difference(self, services_old: list[dict], services_new: list[dict]):
        services_old_dict = {service["Name"]: service for service in services_old}
        services_diff = []
        if len(services_new) != len(services_old):
            raise ValueError(f"Different number of services. "
                             f"Current: {len(services_new)}, Backup: {len(services_old)}")
        for serv_cur in services_new:
            serv_name = serv_cur['Name']
            serv_bac = services_old_dict[serv_name]
            if serv_cur != serv_bac:
                serv_diff = {key: {"OldValue": serv_bac[key], "NewValue": serv_cur[key]}
                             for key in self._only_save_rows if serv_cur[key] != serv_bac[key]}
                services_diff.append(serv_diff)
        return services_diff

    def print_backup_difference(self, filename) -> list[dict]:
        services_backup = self.load_services_from_file(filename)
        services_backup_dict = {service["Name"]: service for service in services_backup}
        services_current = self.list_services()
        services_diff = []
        if len(services_current) != len(services_backup):
            raise ValueError(f"Different number of services. "
                             f"Current: {len(services_current)}, Backup: {len(services_backup)}")
        for serv_cur in services_current:
            serv_name = serv_cur['Name']
            serv_bac = services_backup_dict[serv_name]
            if serv_cur != serv_bac:
                serv_diff = serv_cur.copy()
                serv_diff.update({key: f"{serv_cur[key]} -> {serv_bac[key]}"
                             for key in serv_cur.keys() if serv_cur[key] != serv_bac[key]})
                services_diff.append(serv_diff)
        print(json.dumps(services_diff, indent=4, ensure_ascii=False))
        return services_diff


@pyuac.main_requires_admin
def main():
    prog = Program()
    prog.print_services()
    # prog.backup_services("services.json")
    prog.print_backup_difference("services.json")
    pass

if __name__ == '__main__':
    main()
