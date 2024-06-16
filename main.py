import re
from typing import List
from urllib.request import urlopen
import html
from race import Race
import dateparser
import xlsxwriter

# Constantes à modifier à convenance
ANNEE_EN_COURS = "2024"
QUERIED_CATEGORY = "Access"  # Access, Open, Elite

# Constantes à ne pas modifier
MAIN_URL = "https://velopressecollection.ouest-france.fr"
CALENDAR_URL = MAIN_URL + "/route/calendrier/"
MAIN_URL_REGEX = "[0-9]*-[a-z]*-%s-calendrier-des-courses-cyclistes-sur-route.html" % ANNEE_EN_COURS
RACE_DATA_REGEX = "<td>([\s\S]*?)<\/td>"
MONTH_RACE_REGEX = "<tr>([\s\S]*?)<\/tr>"
ACTUALITES_HTML_REGEX = "\/actualites\/.*\.html"
MEDIA_HTML_REGEX = "\/media\/.*\.jpg"
FILTER_PLACE_REGEX = "[A-ZÉÈÇÖËÊÂÄ\-\’\'\s]{2,1000}"
WEBSITE_SEPARATOR = "*******"

if __name__ == "__main__":

    # Fetch data from MAIN_URL
    main_url = CALENDAR_URL
    month_response = urlopen(main_url)
    month_html_bytes = month_response.read()
    month_html = html.unescape(month_html_bytes.decode("utf-8"))

    links_array = []
    all_links = re.findall(MAIN_URL_REGEX, month_html)
    for link in all_links:
        if link not in links_array:
            links_array.append(link)

    races_array: List[Race] = []
    table_header = "<table border=\"1\" cellpadding=\"0\" cellspacing=\"0\">"
    table_footer = "</table>"

    for link in links_array:
        month_response = urlopen(main_url + link)
        month_html_bytes = month_response.read()
        month_html = html.unescape(month_html_bytes.decode("utf-8"))

        month_table_start_index = month_html.find(table_header) + len(table_header)
        month_table_end_index = month_html.find(table_footer)
        month_table = month_html[month_table_start_index:month_table_end_index]

        rows_content = re.findall(MONTH_RACE_REGEX, month_table)
        for row in rows_content:
            rows_content = re.findall(RACE_DATA_REGEX, row)
            filtered_row_content = []
            for row_data in rows_content:
                if len(re.findall(ACTUALITES_HTML_REGEX, row_data)) > 0:
                    row_data = re.findall(ACTUALITES_HTML_REGEX, row_data)[0]
                if len(re.findall(MEDIA_HTML_REGEX, row_data)) > 0:
                    row_data = re.findall(MEDIA_HTML_REGEX, row_data)[0]
                filtered_row_content.append(row_data
                                            .replace("\n", ' ').replace("\t", '')
                                            .replace("<br />", '')
                                            .replace("<p>", '').replace("</p>", '')
                                            .replace("<strong>", '').replace("</strong>", '')
                                            .replace("<b>", '').replace("</b>", '')
                                            )
            if filtered_row_content[0] != WEBSITE_SEPARATOR and QUERIED_CATEGORY in filtered_row_content[2]:
                race: Race = Race()
                french_date_to_date = dateparser.parse(filtered_row_content[0])
                if french_date_to_date is not None:
                    race.date = french_date_to_date
                race.place = re.findall(FILTER_PLACE_REGEX, filtered_row_content[1])[0]
                race.category = filtered_row_content[2]
                race.race_name = filtered_row_content[3]

                # Résoudre le cas où il n'y a pas d'infos
                if len(filtered_row_content) == 6:
                    if filtered_row_content[5] is not None:
                        race.cycling_club = filtered_row_content[4]
                        race.department = filtered_row_content[5]
                else:
                    if filtered_row_content[4] != ' ':
                        race.info_link = MAIN_URL + filtered_row_content[4]
                    race.cycling_club = filtered_row_content[5]
                    if filtered_row_content[6].isdigit():
                        race.department = int(filtered_row_content[6])
                races_array.append(race)

    # Create Excel doc
    workbook = xlsxwriter.Workbook('Courses_' + QUERIED_CATEGORY + '_' + ANNEE_EN_COURS + '.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.set_column(0, 0, 13)
    worksheet.set_column(1, 1, 4)
    worksheet.set_column(2, 2, 43)
    worksheet.set_column(3, 3, 63)
    worksheet.set_column(4, 4, 15)

    worksheet.write(0, 0, "Date")
    worksheet.write(0, 1, "Département")
    worksheet.write(0, 2, "Lieu")
    worksheet.write(0, 3, "Catégorie")
    worksheet.write(0, 4, "Informations")

    worksheet.set_row(0, cell_format=workbook.add_format({'bg_color': "#b0b0b0"}))

    colors = [
        workbook.add_format({'bg_color': "#b3d9ff", 'border': 1}),
        workbook.add_format({'bg_color': "#f08080", 'border': 1}),
        workbook.add_format({'bg_color': "#93e9be", 'border': 1})]

    last_color = colors[0]

    row = 1
    for race in races_array:
        # Column de la date
        if race.date is not None:
            last_color = colors[race.date.month % len(colors)]
            worksheet.write(row, 0, race.date.strftime("%d-%m-%Y"), last_color)
        else:
            worksheet.write(row, 0, "", last_color)

        # Column du département
        if race.department is not None:
            worksheet.write(row, 1, int(race.department), last_color)
        else:
            worksheet.write(row, 1, "", last_color)

        # Column du lieu
        worksheet.write(row, 2, race.place, last_color)

        # Column de la catégorie
        worksheet.write(row, 3, race.category, last_color)

        # Column des infos
        if race.info_link is not None:
            worksheet.write_url(row, 4, race.info_link, string="Lien infos", cell_format=last_color)
        else:
            worksheet.write(row, 4, "", last_color)

        row += 1

    worksheet.autofilter(0, 1, 200, 1)
    workbook.close()
