def log_parser(log_file_name):
    #  Возвращает текст из лог-файла
    text_list = []
    with open(log_file_name) as log_file:
        # list = [row.strip().split("\n") for row in log_file]
        for row in log_file:
            text_list.append(row)
    text_list.reverse()
    return ' '.join(text_list)


def get_config(conf):
    #  Парсит конфиг
    with open(conf) as config:
        array = [row.strip().split("|") for row in config]
    return array
