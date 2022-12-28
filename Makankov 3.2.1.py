import csv
import os
import re
from os import path
from typing import List, Dict


class SplitCsvFileByYear:
    """
    Класс для раделения набора вакансий по годам
    Attributes:
        file_name: Название файла
        :type (str)
        dir_name: Название папки, в которой хранятся итоговые csv-файлы
        :type (List[Vacancy])
        headlines: Названия загаловков
        :type (List[str])
        vacancies: Набор вакансий
        :type (List[List[str]])
    """
    def __init__(self, file_name : str, directory : str):
        """
        @param file_name: Название файла
        :type (str)
        @param file_name: Название папки, в которой хранятся итоговые csv-файлы
        :type (str)
        """
        self.file_name = file_name
        self.dir_name = directory
        self.headlines, self.vacancies = self.__csv_reader()
        self.__csv_process(self.headlines, self.vacancies)

    def __csv_reader(self) -> (List[str], List[List[str]]):
        """
        Читает из csv файла вакансии и возвращает в виде списка загаловков и набора вакансий
        @return: Список загаловков и набора вакансий
        :type (List[str], List[List[str]])
        """
        with open(self.file_name, encoding='utf-8-sig') as file:
            file_reader = csv.reader(file)
            lines = [row for row in file_reader]
            headlines, vacancies = lines[0], lines[1:]
        return headlines, vacancies

    def __csv_process(self, headlines : List[str], vacancies : List[List[str]]) -> None:
        """
        Обрабатывает полученный набор вакансий и загаловков
        @param headlines: Названия загаловков
        :type (List[str])
        @param vacancies: Набор вакансий
        :type (List[List[str]])
        @return: None
        """
        cur_year = "0"
        os.mkdir(self.dir_name)
        vacancies_cur_year = []
        for vacancy in vacancies:
            if (len(vacancy) == len(headlines)) and (all([v != "" for v in vacancy])):
                vacancy = [" ".join(re.sub("<.*?>", "", value).replace('\n', '; ').split()) for value in vacancy]
                vacancy_dict = {x: y for x, y in zip([r for r in headlines], [v for v in vacancy])}
                if vacancy[-1][:4] != cur_year:
                    if len(vacancies_cur_year) != 0:
                        self.__csv_writer(headlines, vacancies_cur_year, cur_year)
                        vacancies_cur_year.clear()
                    cur_year = vacancy[-1][:4]
                vacancies_cur_year.append(vacancy_dict)
        self.__csv_writer(headlines, vacancies_cur_year, cur_year)

    def __csv_writer(self, headlines : List[str], vacancies : List[Dict[str, str]], cur_year : str) -> None:
        """
        Записывает данные в csv-файл
        @param headlines: Названия загаловков
        :type (List[str])
        @param vacancies: Набор вакансий
        :type (List[List[str]])
        @param cur_year: Текущий год обработки
        :type (str)
        @return: None
        """
        name = path.splitext(self.file_name)
        with open(f'{self.dir_name}/{name[0]}_{cur_year}.csv', 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=headlines)
            writer.writeheader()
            writer.writerows(vacancies)


print("Введите название файла: ", end="")
SplitCsvFileByYear(input(), " vacancies_by_year")