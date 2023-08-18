@ECHO OFF
ECHO Executing authentication
curl -k -X POST --data "@mycreds.json" -H "Content-Type: application/json" -c COOKIE.txt https://10.40.145.1/api/aaaLogin.json
PAUSE



ECHO ################# Hooray All scripts has been executed.###############