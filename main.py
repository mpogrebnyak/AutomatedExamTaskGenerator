import pandas as pd
from generate_individual_tasks import GenerateIndividualTasks
from generate_exam_tickets import GenerateExamTickets


CONFIG_FILE = "config.xlsx"


def read_config():

    settings = pd.read_excel("config.xlsx", sheet_name="settings")
    topics = pd.read_excel("config.xlsx", sheet_name="topics")
    text = pd.read_excel("config.xlsx", sheet_name="text")

    config = dict(zip(settings["key"], settings["value"]))
    config["questions_per_topic"] = dict(zip(topics["topic"], topics["count"]))
    config["text"] = dict(zip(text["key"], text["value"]))

    return config


def main():

    config = read_config()

    mode = str(config["mode"]).lower()

    if mode == "tasks":
        generate_individual_tasks = GenerateIndividualTasks(config)
        generate_individual_tasks.generate()

    elif mode == "exam":
        generate_exam_tickets = GenerateExamTickets(config)
        generate_exam_tickets.generate()

    else:
        raise ValueError("mode must be 'tasks' or 'exam'")


if __name__ == "__main__":
    main()