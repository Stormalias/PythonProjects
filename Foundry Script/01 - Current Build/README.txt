This README assumes you know why you are using this script. For more verbose description please refer to the Knowledge Base Page assigned to this

1. Setup python modules needed via pip and requirements.txt. 
!!!!!!!!!!MAKE SURE TO CD INTO DIRECTORY OR PROVIDE THE ABSOLUTE PATH TO REQUIREMENTS.TXT!!!!!!!!!!!!!!!!!!!!!!!

	> Open command prompt
	> $pip install - r requirements.txt

2. Install ODBC driver from this link:
https://lava.palantircloud.com/workspace/documentation/product/foundry-bi-tool-integration/downloads#odbc-driver

Download and install Foundry ODBC Driver

!!!!!!!!!!you're looking for ODBC. You are NOT LOOKING for JDBC!!!!!!!!!!!!!!!!!!!!!!!!!!!!
3. Generate token:
	> Go to  https://lava.palantircloud.com/workspace/settings/tokens
	> Click on 'create token' at top right
	> Enter any name and description. Pick maximum validity available (usually 2 weeks max)
	> Click Generate
	!!!!!!!!!!! VERY IMPORTANT. AFTER YOU GENERATE IF YOU CLICK OFF THE GIVEN TOKEN OR CLICK OK YOU WILL NOT BE ABLE TO SEE THE TOKEN AGAIN!!!!!!!!!!!!!
	> Copy token to clipboard for later use.
	
4. From the start menu look for "Windows Administrative Tools"

5. From the list of programs in the folder, find and open "ODBC Data Sources (64 bit)"

6. Add data source:
	> Click add
	> Pick FoundrySqlDriver and click finish
	> Config items:
		1. Data Source Name - Enter any name. Recomended is 'foundry'. If you choose another name remember to change 
							  the name in csvExtractor.py 
		2. Description		- Give any description you want
		3. Server			- https://lava.palantircloud.com
		4. Token			- paste in token from previous step
	> Click 'test' to make sure configuration is good.

7. Run FoundryMonitoringScript and enjoy. You will need to regenerate a new token every 2 weeks and replace it in the data source cofig. Otherwise
the only other change you will make will be the name of the Data Source in csvExtractor.py