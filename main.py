import argparse
import copy
import datetime
import json
import locale
import pathlib
import re
import subprocess
import sys
import simplejson
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

    def _print_services(self):
        services = self.list_services()
        msg = json.dumps(services, indent=4, ensure_ascii=False)
        # split = msg.splitlines()
        # for line in split:
        #     print(line)
        print(msg)

    def list_services(self) -> list[dict]:
        cmd = """
        [Console]::OutputEncoding = [Text.Encoding]::UTF8
        Get-Service | Select-Object -Property Name,DisplayName,Status,StartType | ConvertTo-Json -EnumsAsStrings
        """

        # result = subprocess.getoutput(f"pwsh -NoProfile -Command \"{cmd}\"")
        completed = subprocess.run(["pwsh", "-NoProfile", "-Command", cmd],
                                   capture_output=True)
        if completed.returncode != 0:
            print("Failed to get some services:")
            print(bytes(completed.stderr).decode('utf-8'))
        try:
            services_json_str = bytes(completed.stdout)
            # print(f"LOCALE={locale.getlocale()}")
            # lines = services_json_str.splitlines()
            # first_line_num_not_ansi = next(i for i, line in enumerate(lines) if any(c for c in line if not 0 <= c <= 127))
            # print("RAW")
            # print(*lines[first_line_num_not_ansi:first_line_num_not_ansi+10], sep='\n')
            # decoding_format = 'utf-8'
            # print(f"{decoding_format.upper()} DECODED")
            # coded = bytes(completed.stdout).decode('utf-8')
            # coded_part = '\n'.join(coded.splitlines()[first_line_num_not_ansi:first_line_num_not_ansi+10])
            # print(coded_part)
            # print(services_json_str)
            services_json_str_decoded = services_json_str.decode("utf-8")
            services: list[dict] = json.loads(services_json_str_decoded)
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

    def print_backup_difference(self, filename, ignore_suffix=False) -> None:
        services_backup = self.load_services_from_file(filename)
        services_backup_dict = {service["Name"]: service for service in services_backup}
        services_current = self.list_services()
        services_current_dict = {service["Name"]: service for service in services_current}
        if ignore_suffix:
            for services_dict in [services_backup_dict, services_current_dict]:
                new_dict = dict()
                for service_name in services_dict.keys():
                    new_name = re.sub(r"_[a-z0-9]{,10}$", "", service_name)
                    new_service = services_dict[service_name].copy()
                    new_service["Name"] = new_name
                    new_dict[new_name] = new_service
                services_dict.clear()
                services_dict.update(new_dict)
        services_diff = []
        # if len(services_current) != len(services_backup):
        #     raise ValueError(f"Different number of services. "
        #                      f"Current: {len(services_current)}, Backup: {len(services_backup)}")
        for serv_cur in services_current_dict.values():
            serv_name = serv_cur['Name']
            if serv_name not in services_backup_dict:
                continue
            serv_bac = services_backup_dict[serv_name]
            if serv_cur != serv_bac:
                if serv_cur["StartType"] == serv_bac["StartType"]:
                    continue
                serv_diff = copy.deepcopy(serv_cur)
                serv_diff.update({key: f"{serv_bac[key]} -> {serv_cur[key]}"
                             for key in serv_cur.keys() if serv_cur[key] != serv_bac[key]})
                services_diff.append(serv_diff)
        if len(services_diff) > 0:
            print(json.dumps(services_diff, indent=4, ensure_ascii=False))
        deleted_services = [service for service in services_backup_dict.values()
                            if service["Name"] not in services_current_dict]
        if len(deleted_services) > 0:
            print(f"Deleted services:")
            print(json.dumps(deleted_services, indent=4, ensure_ascii=False))
        new_services = [service for service in services_current_dict.values()
                        if service["Name"] not in services_backup_dict]
        if len(new_services) > 0:
            print(f"New services:")
            print(json.dumps(new_services, indent=4, ensure_ascii=False))

    def _save_services(self):
        date_str = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        filename = pathlib.Path(f'services_{date_str}.json')
        self.backup_services(filename)
        print(f"Services saved to \"{filename.absolute()}\"")
        pass

    def _print_diff(self, filename, ignore_suffix = False):
        if not pathlib.Path(filename).is_file():
            print(f"File \"{filename}\" does not exist")
            sys.exit(1)
        self.print_backup_difference(filename, ignore_suffix)
        pass

    @classmethod
    def main(cls, args=None):
        prog = Program()
        if args is None:
            args = sys.argv[1:]
        parser = argparse.ArgumentParser(
            prog='ServicesBackup',
            description='Backup and restore services start states',
            epilog='Created by MrMarvel')
        subparsers = parser.add_subparsers(help='sub-command help')
        parser_save = subparsers.add_parser('save', help='save help')
        parser_save.set_defaults(func=lambda _: prog._save_services())
        parser_print = subparsers.add_parser('print', help='print help')
        parser_print.set_defaults(func=lambda _: prog._print_services())
        parser_diff = subparsers.add_parser('diff', help='print help')
        parser_diff.add_argument('filename', type=str, help='File to compare with')
        parser_diff.add_argument('--ignore-suffix', help='Ignore suffix in service name',
                                 action='store_true', default=False, dest='ignore_suffix')
        parser_diff.set_defaults(func=lambda x: prog._print_diff(x.filename, x.ignore_suffix))
        parsed = parser.parse_args(args)
        parsed.func(parsed)


@pyuac.main_requires_admin
def main():
    Program.main()
    # prog.print_services()
    # # prog.backup_services("services.json")
    # prog.print_backup_difference("services.json")
    pass

if __name__ == '__main__':
    main()
