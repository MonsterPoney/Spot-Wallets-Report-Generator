# Spot-Wallets-Report-Generator
Daily report generator for spot wallets (Binance, Bybit and Kucoin), keep in mind that prices are an average given by the API.
Start it more than once a day will rewrite previous daily data.

The program use EPPlus for the worksheet generation, more info on it here : [EPPlus Software](https://www.epplussoftware.com/)

If you build the source, change the property 'Copy to Output Directory' of the config.ini to 'Do not copy' after the first build or it will change the encoding everytime.
Then change the encoding to UTF-8 with notepad++ or other text editor.

After configuration, execute 'Spot Wallets Report Generator.exe' in the Release folder to start the program.

## Configuration

### Binance
Acquire Binance API key in your profile > API Management
Edit restrictions to check "Enable Reading" only, unckeck everything else.
You can also restrict the access to one or multiple IPs.

![Binance API config](https://user-images.githubusercontent.com/25821500/154150789-b6f87351-493d-400b-89fa-a2c4af0ef699.JPG)

Copy API key and secret Key to paste them in the config.ini file.
Switch 'UseBinance' to true.

### Bybit
Acquire Bybit API key in your profile > API
In the "API key permissions" section, check "Read-Only" and check the permission "Trade" for SPOT
You can also restrict the access to one or multiple IPs.

![Bybit API config](https://user-images.githubusercontent.com/25821500/154150876-78eb9950-defe-4912-b8d1-c521002e5e96.JPG)

Copy API key and secret Key to paste them in the config.ini file.
Switch 'UseBybit' to true.

### Kucoin
Acquire KuCoin API key in your profile > API Management
Set a passphrase
Restrict the API to General only
You can also restrict the access to one or multiple IPs.

![Kucoin API config](https://user-images.githubusercontent.com/25821500/154150930-09cb7f74-972d-41f4-b9fc-a0f408103262.JPG)

Copy API key, secret Key and passphrase to paste them in the config.ini file.
Switch 'UseKucoin' to true.

### Configuration file
This file is mainly used for APIs credentials.
But also :
- Specify arguments.
	You can also execute the program with the cmd and specify arguments.

- Choose to store the data in a on-disk database file.

- Use an evolution chart of BTC or/and USDT.

- Ignore asset with a total value under a certain amount of USDT.

- Sort the report by asset name or by plateform. 

## Disclamer
Even if your API credentials are linked to read-only endpoints, I advise NOT to share the config file.
The data used by this program is contain in the .db file, you can refuse to store data by switching the "Database" option to "False" in the config.ini file.
No data is sent on an external source, everything is on your computer in the program folder per default.

## Troubleshoot
- Error "database is locked"
	This happen when a connection to the db file is already openned
	
- Error code 1398 or -1021 "Timestamp for this request was 1000ms ahead of the server's time."
	Your system clock might be (at least slightly) off, try resync the system clock of your computer in the Date and time settings.
	The program will attempt to do it if you set the 'AutoTimeSync' to 'true'.

- Error when reading the 'xxxx' option
	If you build the source code, visual studio encode the config file with a Byte Of Mark (UTF-8-BOM).
	Change the encoding to UTF-8 with notepad++ or other text editor.

- Error code 1610
	Check the config.ini file, there must be a missing parameter like a boolean.

- Error code 1065
	Error with the database, the log should give an explanation

- Error code 259
	No asset was recovered, make sure you use at least one API in the config file.

- Error code 574
	It's the global error, ( ._.) That's annoying but the log file should help.
	
## Donations
I developped this program on my spare time, if you want to support me you can donate at theses adresses :
- BTC : bc1q3u5m3xq66gu57cf2v25rkr0qwt3v9evtnzuej7
- ETH, USDT(ERC-20) & smart chain : 0x356D763b6924D7DC864c550941B911ca87a98e26
- BNB : bnb1rwgj40gyesrnszwvfgq5kgttet654jynv0ql7r
