import json
from pathlib import Path
import sys

import requests  # pip install requests
import xlwings as xw  # pip install xlwings


def main():
    wb = xw.Book.caller()
    sht = wb.sheets[0]

    # Get value from Name Range
    city_name = sht.range("city_name").value

    # Get City ID
    URL_CITY = f"https://www.metaweather.com/api/location/search/?query={city_name}"
    response_city = requests.request("GET", URL_CITY)
    try:
        city_title = json.loads(response_city.text)[0]["title"]
    except IndexError:
        msgbox_invalid_city = wb.macro("InvalidCity")
        msgbox_invalid_city(city_name)
        sys.exit()
    city_id = json.loads(response_city.text)[0]["woeid"]

    # Get Weather for City ID
    URL_WEATHER = f"https://www.metaweather.com/api/location/{city_id}/"
    response_weather = requests.request("GET", URL_WEATHER)
    weather_data = json.loads(response_weather.text)["consolidated_weather"]

    # Create empty lists
    min_temp = []
    max_temp = []
    weather_state_name = []
    weather_state_abbr = []
    applicable_date = []

    # Iterate over weather_data & append data to lists
    for index, day in enumerate(weather_data):
        min_temp.append(weather_data[index]["min_temp"])
        max_temp.append(weather_data[index]["max_temp"])
        weather_state_name.append(weather_data[index]["weather_state_name"])
        weather_state_abbr.append(weather_data[index]["weather_state_abbr"])
        applicable_date.append(weather_data[index]["applicable_date"])

    # Return Weather Forecast back to Excel
    sht.range("C5").value = applicable_date
    sht.range("C6").value = weather_state_name
    sht.range("C7").value = max_temp
    sht.range("C8").value = min_temp
    sht.range("D3").value = city_title

    # Create List
    icon_names = ["one_day", "two_day", "three_day", "four_day", "five_day", "six_day"]

    # Changed Path.cwd() to Path(__file__).parent
    icon_path = Path(__file__).parent / "images"

    # Iterate over icon_name & weather_state_abbr
    # Update images of workbook
    for icon, abbr in zip(icon_names, weather_state_abbr):
        image_path = Path(icon_path, abbr + ".png")
        sht.pictures.add(image_path, name=icon, update=True)


if __name__ == "__main__":
    xw.Book("weatherapp.xlsm").set_mock_caller()
    main()
