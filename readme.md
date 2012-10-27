Reference: https://www.virustotal.com/documentation/public-api/

All public functions of this module correspond to actions available 
through the public VirusTotal API v.2. All interactions with the API require an
API key. 

The Virus Total API returns JSON. Luckily, REALstudio has shipped with built-in 
JSON support since RS2011r2. All the public functions of this module, therefore, 
return JSONItems (or Nil, on error.)

If the returned JSONItem was not Nil, then LastResponseCode and 
LastResponseVerbose correspond to the response_code and verbose_msg members 
of Virus Total's response. 

If the returned JSONItem was Nil, and it was because of a socket error, then 
LastResponseCode and LastResponseVerbose correspond to the RB socket error 
number and a brief error message. 

If the returned JSONItem was Nil, and it was because of a JSON error, then 
LastResponseCode is VTAPI.INVALID_RESPONSE and LastResponseVerbose is a 
brief error message. 