import os
from collections import Counter

import pandas as pd


def list_to_upper_case(data_list: list) -> list:
    """
    Функция для изменения регистра названий в списке
    :param data_list: список с данными
    :return: возвращает новый список с элементами в новом регистре
    """
    return_list = []
    for i in data_list:
        return_list.append(i.upper())
    return return_list


def main():
    os.chdir("excel/input")
    input_xlsx_df = pd.read_excel("input.xls")['WoS Categories']
    oecd_category_df = pd.read_excel("OECD Category Mapping.xlsx")
    input_xlsx_list: list = input_xlsx_df.to_list()
    input_xlsx_df_upper: list = list_to_upper_case(input_xlsx_list)

    result_list: list = []
    get_result_oecd_and_wos_category: list = []
    get_list_all_wos_category: list = []

    uniq_dict_description = dict.fromkeys(oecd_category_df["Description"].to_list(), "")
    list_description = oecd_category_df["Description"].to_list()
    list_wos_description = oecd_category_df["WoS_Description"].to_list()

    for i in range(len(list_description)):
        uniq_dict_description[list_description[i]] += list_wos_description[i] + ";"

    for elem in input_xlsx_df_upper:
        dict_append = {}
        for key, item in uniq_dict_description.items():
            for i in elem.split("; "):
                if item.find(i) != -1:
                    dict_append |= {key: i}
        get_result_oecd_and_wos_category.append(dict_append)

    for element in get_result_oecd_and_wos_category:
        append_list = []
        for item in element.values():
            append_list.append(item)
        get_list_all_wos_category.append(list(set(append_list)))

    for item in get_list_all_wos_category:
        for values in item:
            result_list.append(values)

    dict_title_value = dict(Counter(result_list))
    list_to_df_title = list(dict_title_value.keys())
    list_to_df_category = list(dict_title_value.values())

    title_category_df = pd.DataFrame({"wos_category": list_to_df_title, "value": list_to_df_category})
    result_df = pd.merge(left=oecd_category_df, right=title_category_df, right_on="wos_category",
                         left_on="WoS_Description")
    os.chdir("../output")
    result_df.to_excel("result_new.xlsx", index=False)
