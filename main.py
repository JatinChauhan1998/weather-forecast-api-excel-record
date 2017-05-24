import pyowm
import xlsxwriter

workbook = xlsxwriter.Workbook('sample-agriculture-record.xlsx')
worksheet = workbook.add_worksheet()

owm=pyowm.OWM('3b07f45dfa53ff9aa2956b9d7ba13d12')

tomorrow=pyowm.timeutils.tomorrow()

observation = owm.weather_at_place("Chandigarh,IN")
w=observation.get_weather()
weather_forecast=str(w)
weather_humidity=str(w.get_humidity())
weather_temperature=str(w.get_temperature('celsius'))
weather_wind=str(w.get_wind())

weather = (
    ['Weather Forecast', weather_forecast],
    ['Wind', weather_wind],
    ['Humidity',  weather_humidity],
    ['Temperature',  weather_temperature],
)

row = 0
col = 0

for name, value in (weather):
    worksheet.write(row, col, name)
    worksheet.write(row, col + 1, value)
    row += 1

workbook.close()
