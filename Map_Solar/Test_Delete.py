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

# results = solarwinds.query (    """
#                                 SELECT
#                                     EventID,
#                                     EventTime,
#                                     EventType,
#                                     TimeStamp,
#                                     DisplayName,
#                                     Description,
#                                     InstanceSiteId,
#                                     Message
#                                 FROM
#                                     Orion.Events
#                                 """
#                             )

results = solarwinds.query (    """
                                SELECT
                                    n.Name,
                                    n.DisplayName,
                                    n.Description,
                                    m.Name,
                                    m.FullName,
                                    m.Description
                                FROM
                                    Orion.Container n
                                INNER JOIN
                                    Orion.ContainerMembers m
                                ON
                                    n.ContainerID=m.ContainerID
                                """
                            )

print ( results )