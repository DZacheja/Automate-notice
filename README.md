# Automate-notice
This Word file with VBA macrosa generate douments. It's very helpful in notice on land plots in geodesy.

@User manual:

1. Create a word file with Bookmarks:
-Name
-Street
-Postcode
-ID
-contents

!You must have the same Bookmarks name as above!

2. Create Excel File witch contains on columns:
*Sheet 1 
	-Land plot number

	-Land plots date (numbers of days after start date, eg.: when you setup start date to
	01.01.2020 and you want create description for 05.01.2020 you have to insert text "4", because
	day 01.01.2020 + 4days = 05.01.2020 - that format its necessery if one land plot have more
	than one date eg.: 1,3,5 ..)

	-Land plots hour (program adds automatically suffix ":00")

*Sheet 2 
	-Person name
	-Person Street
	-Person Postcode
	-Person ID
	-Person LandPlots(Like a date and hour you can write more than one land plot.
			  If you want add more plots, just separate them with "," eg.: 100/1,100/2,100/3)
3. Open DOCUMENT_GENERATOR and load excel, and word files
4. Fill in the columns index for all information
5. Create the descrition whitch will be put to the contents Bookmarks.
6. Now click OK and wait 

