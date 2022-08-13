import os
import sys
import requests
import json
from time import sleep
import configparser
from win32com.client import Dispatch


class Config: #TODO: use python configparser and update the set and write method
    def __init__(self, path):
        self.path = path
        self.config = configparser.ConfigParser()
        self.read()

    def read(self):
        self.config.read(self.path)
        return self.config

    def write(self):
        with open(self.path, "w", encoding="utf-8") as configfile:
            self.config.write(configfile, space_around_delimiters=False)

    def has_section(self, section):
        return self.config.has_section(section)

    def has_option(self, section, option):
        return self.config.has_option(section, option)

    def add_section(self, section):
        self.config.add_section(section)
        self.write()

    def set(self, section, option, value):
        if not self.has_section(section):
            self.add_section(section)
        self.config.set(section, option, value)
        self.write()

    def get(self, section, option):
        return self.config.get(section, option)


def write_file(filename, data):
    with open(filename, "w", encoding="utf-8") as f:
        f.write(data)


def read_file(filename):
    with open(filename, "r", encoding="utf-8") as f:
        return int(f.read())


def get_latest_build():
    try:
        response = requests.get(PAPER_URL)
        json_data = json.loads(response.text)
        return (str(json_data["builds"][-1]), json_data)
    except Exception as e:
        raise Exception(f"Invalid URL or some error occured while making the GET request to the specified URL:\n{e}")


def create_shortcut(path, target="", w_dir="", icon="", arguments=""):
    shell = Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.Arguments = f"\"{arguments}\""
    shortcut.WorkingDirectory = w_dir
    shortcut.WindowStyle = 7
    if icon != "":
        shortcut.IconLocation = icon
    shortcut.save()


"""This part creates if it doesn't exist the needed directories and make sure the config.ini file exists in the user's home directory."""
temp_dir = os.path.dirname(os.path.abspath(os.path.realpath(__file__)))
file_path = os.path.abspath(os.path.realpath(sys.argv[0])) # file path of the current file
file_dir = os.path.dirname(file_path) # directory of the current file
appdata_dir = os.getenv("APPDATA")
startup_dir = os.path.join(appdata_dir, r"Microsoft\Windows\Start Menu\Programs\Startup")
home_dir = os.path.expanduser("~")
thegeeking_dir = os.path.join(home_dir, ".TheGeeKing")
webhook_paper_notifier_dir = os.path.join(thegeeking_dir, "Paper-Notifier")
# create a directory called .TheGeeKing if it doesn't exist
if not os.path.exists(thegeeking_dir): os.mkdir(thegeeking_dir)
# create a directory called Discord-RPC-Maker if it doesn't exist
if not os.path.exists(webhook_paper_notifier_dir): os.mkdir(webhook_paper_notifier_dir)
# if config.ini doesn't exist, create it and write the default values
if not os.path.exists(os.path.join(webhook_paper_notifier_dir, "config.ini")):
    with open(os.path.join(webhook_paper_notifier_dir, "config.ini"), "w", encoding="utf-8") as f:
        f.write("")
config = Config(os.path.join(webhook_paper_notifier_dir, "config.ini"))
if not config.has_section("config"):
    config.set("config", "MINECRAFT_VERSION", "1.19.2")
    config.set("config", "WEBHOOK_URL", "Your Discord Webhook URL.")
    config.set("config", "CHECK_EVERY", str(60*60*24)) # 1 day

if not config.get("config", "WEBHOOK_URL").startswith("https://discord.com/api/webhooks/"):
    os.system(f"start {os.path.join(webhook_paper_notifier_dir, 'config.ini')}")
    os._exit(0)

VERSION = config.get("config", "MINECRAFT_VERSION")
WEBHOOK_URL = config.get("config", "WEBHOOK_URL")
PAPER_URL = f"https://api.papermc.io/v2/projects/paper/versions/{VERSION}"
CHECK_EVERY = int(config.get("config", "CHECK_EVERY"))

if not config.has_option("config", "LATEST_BUILD"):
    config.set("config", "LATEST_BUILD", get_latest_build()[0])

create_shortcut(os.path.join(startup_dir, "Paper-Notifier - Startup.lnk"), file_path, file_dir, os.path.join(temp_dir, "MMA.ico"))

while True:
    current_build = config.get("config", "LATEST_BUILD")
    latest_build, json_data = get_latest_build()
    if latest_build > current_build:
        dl_link = f"https://api.papermc.io/v2/projects/paper/versions/{json_data['version']}/builds/{latest_build}/downloads/paper-{json_data['version']}-{latest_build}.jar"
        config.set("config", "LATEST_BUILD", latest_build)
        #for all params, see https://discordapp.com/developers/docs/resources/webhook#execute-webhook
        data = {
            "content": None,
            "embeds": [
                {
                    "title": "Honey wake up new paper build just dropped",
                    "description": f"New version is {json_data['version']}-{latest_build}!",
                    "url": dl_link,
                "color": 4408131,
                "author": {
                    "name": "MMA | TheGeeKing",
                    "url": "https://github.com/TheGeeKing",
                    "icon_url": "https://avatars.githubusercontent.com/u/58857539?v=4"
                    }
                }
                ]
            }

        result = requests.post(WEBHOOK_URL, json=data)
    sleep(CHECK_EVERY)
