@echo off
SETLOCAL EnableDelayedExpansion

echo.
echo Please select your option (1 or 2)
echo.
echo 1. Validate BOM Template Document
echo 2. Generate Assent Template
set /p userInput=""

echo.

if %userInput% ==1 (
 	echo You have selected the option to 'Validate BOM template'

	echo.
  
	echo Please enter the MSTR excel file name along with extension
	set /p mstrFileName=""
	
	echo.
  
	echo Please enter the Commodity excel file name along with extension
	set /p commodityFileName=""
	
) else if %userInput% ==2 (
  	echo You have selected the option to 'Generate Assent template'
  	
  	echo.
  
	echo Please enter the Assent template -System generated- file name along with extension
	set /p assentTemplateFileName=""
	
	echo.
	
	echo "Please enter the customer name"
	set /p customerName=""

	echo.	
		
	echo "Please enter the Ticket number"
	set /p ticketNumber=""
	
) else (
	echo Invalid value... Enter 1 or 2 ..
	echo Bailing out...
    goto:eof
)

echo.

echo Please enter the BOM Template excel file name (along with extension) 
set bomTemplateFileName=
set /p bomTemplateFileName=""
echo.

echo "Please wait for the process to complete..."

echo.
echo.

java -Xmx1024m -jar -DuserInput="%userInput%" -DcommodityFileName="%commodityFileName%" -DbomTemplateFileName="%bomTemplateFileName%" -DmstrFileName="%mstrFileName%" -DassentTemplateFileName="%assentTemplateFileName%" -DcustomerName="%customerName%" -DticketNumber="%ticketNumber%" assenttemplate-0.0.1-SNAPSHOT.jar

cmd /k