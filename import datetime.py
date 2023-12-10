from datetime import datetime
today = datetime.now()
filename = datetime.now().strftime("%d.%m.%Y_%H%M")
print(filename)
#change main 2