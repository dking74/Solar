from orionsdk import SwisClient
import requests

valid = False 

verify = False
if not verify:
    from requests.packages.urllib3.exceptions import InsecureRequestWarning
    requests.packages.urllib3.disable_warnings ( InsecureRequestWarning )


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

 #query = self._solarwinds.query ( "SELECT N.Name, M.Latitude FROM Orion.Groups N INNER JOIN Orion.WorldMap.Point M ON N.ContainerID=M.InstanceID")

results = solarwinds.query (    """
                                SELECT
                                    n.MACCurrentInfo,
                                    m.caption
                                FROM
                                    Orion.Nodes m
                                INNER JOIN
                                    Orion.UDT.MACCurrentInfo n
                                ON
                                    m.NodeID=n.NodeID

                                """
                            )

print ( results )