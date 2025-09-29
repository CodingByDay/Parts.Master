1. Use the following to create the exe file
pyinstaller --onefile --noconsole --clean `
>>   --name PartsMaster `
>>   --icon=assets/official-logo.ico `
>>   --add-data "assets;assets" `
>>   --add-data "app_info.json;." `
>>   main.py
2. Use Inno Setup compiler to open the installer.iss and run the msi creation process