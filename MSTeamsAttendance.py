import pandas as pd
from numpy import timedelta64
from time import sleep

def main():
	print("Welcome to Microsoft Teams Attendance Parser\nPlease make sure the downloaded attendance file is in the same directory as this file.\n")

	# Taking the name of the source
	name = input("Enter your Attendance file name (without .csv):")
	src = (name + ".csv")
	dest = (name + "_converted.csv")

	# make sure file exists
	try:
		f = open(src)
	except IOError:
		print("Fatal Error: " + src + " does not exist.\nExiting in 2 seconds")
		sleep(2)
		exit()
	f.close()

	# Reading the raw data and convering it to a dataframe
	raw = pd.read_csv(src, encoding='utf-16', delimiter='\t')
	data = pd.DataFrame(raw)

	# Converting to pandas timestamp and sorting
	data['Timestamp'] = data['Timestamp'].apply(pd.Timestamp)
	data.sort_values(by='Timestamp', ignore_index=True, inplace=True)

	# Prepping the desired output 
	desired = data.copy()
	desired['Total Time (m)'] = [0 for i in range(desired.shape[0])]
	desired['Still In?'] = [False for i in range(desired.shape[0])]
	desired.rename(columns={'Timestamp':'Last Entry'}, inplace=True,)
	desired.drop_duplicates(subset='Full Name', inplace=True, ignore_index=True)
	desired.drop(columns=['User Action'], inplace=True)

	# Get a histogram to store whether an row has been processed or not
	rows = data.shape[0]
	histogram = [False for i in range(rows)]

	# Main loop
	for i in range(rows):

		# If previously processed then don't process again
		if histogram[i] == True:
			continue
		histogram[i] = True

		if data['User Action'][i] == "Left":
			print("Fatal Error: A \"Left\" found before a \"Joined\"\nExiting in 2 seconds")
			sleep(2)
			exit()

		# Some variables that will be used later
		currentName = data['Full Name'][i]
		currentID = desired.index[desired['Full Name'] == currentName].tolist()[0]
		inRange = False

		# Update timestamp of last entry
		desired.loc[currentID, 'Last Entry'] = data['Timestamp'][i]

		# An attemp to find the leaving entry for that user and updating total time
		for j in range(i + 1, rows):
			if data['Full Name'][j] == currentName:
				if data['User Action'][j] == "Left":
					desired.loc[currentID, 'Total Time (m)'] += (data['Timestamp'][j] - data['Timestamp'][i]) / timedelta64(1,'m')
					inRange = True
					histogram[j] = True
					break
				else:
					print("Fatal Error: Multiple \"Joined\" in Succession\nExiting in 2 seconds")
					sleep(2)
					exit()

		# If an exit isn't found then the user is still in the meeting
		if inRange == False :
			desired.loc[currentID, 'Still In?'] = True


	# Format the output and convert to destination csv file
	desired['Total Time (m)'] = desired['Total Time (m)'].map('{:,.2f}'.format)
	desired['Last Entry'] = desired['Last Entry'].dt.time
	desired.sort_values(by='Full Name', ignore_index=True, inplace=True)
	desired.to_csv(dest, encoding='utf-16', sep='\t', header=True, index=False)


if __name__ == "__main__":
	main()
