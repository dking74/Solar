from orionsdk import SwisClient
import requests

valid = False 

while not valid:

    username = input ( "Username: " )
    password = input ( "Password: " )
    print            (     "\n"     )

    try: 
        solarwinds = SwisClient ( "solarwinds.ameren.com" , username , password )
        solarwinds.query        (     "SELECT AccountID FROM Orion.Accounts"    )
        valid = True

    except requests.exceptions.HTTPError:
        valid = False

results = solarwinds.query (    """
                                SELECT
                                    m.MACAddressInfo.MACAddress,
                                    m.Ports.Name
                                FROM
                                    Orion.Nodes m
                                """
                            )

print ( results )