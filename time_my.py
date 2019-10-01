import datetime

def week_number():
	today = datetime.date.today()
	return (today.strftime("%W"))

def day():
	day=datetime.datetime.today().isoweekday()
	return (day)


def actual_time():
	day=datetime.datetime.today()
	return(day.hour+(day.minute/100))

#print(type(week_number()))