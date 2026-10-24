# Edit by Liam
###Script logic###
# 1.Connect to FortiManager via API
# 2.Get registered devices from different ADOM
# 3.put all devices information are organized in dictionary format and stored in a list


import requests
import urllib3
from pprint import pprint
import os
import time

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def GetAPI(adom_name):
    FMG_IP = os.getenv("FMG_IP", "10.81.253.253")
    FMG_API_KEY = os.getenv("FMG_API_KEY")
    url = f"https://{FMG_IP}:4443/jsonrpc"
    headers = {
        "Authorization": f"Bearer {FMG_API_KEY}",
        "Content-Type": "application/json",
    }

    payload = {
        "id": 1,
        "method": "get",
        "params": [
            {
            "url": f"/dvmdb/adom/{adom_name}/device",
            "fields": ["hostname", "sn", "ip", "platform_str"],
            "option":["no loadsub"]
            },]
            }

    try:
        response = requests.post(
            url,
            headers=headers,
            json=payload,
            verify=False,  # Temporary testing only
            timeout=15
        )

        # Check HTTP response code 200 OK, it will continue.
        # otherwise it will raise requests.exceptions.HTTPError
        response.raise_for_status()  # check HTTP response code,
        result = response.json()

        # Check FortiManager API operation status
        api_result = result["result"][0]
        status = api_result.get("status", {})

        if status.get("code") != 0:
            print(
                f"[Error] ADOM {adom_name} API request failed:"
                f"{status}"
            )
            return None
        return result

    except requests.exceptions.Timeout:
        print(f"[ERROR] ADOM {adom_name}: Request timeout")
    except requests.exceptions.ConnectionError as e:
        print(f"[ERROR] ADOM {adom_name}: Connection failed: {e}")
    except requests.exceptions.HTTPError as e:
        print(f"[ERROR] ADOM {adom_name}: HTTP error: {e}")
    except requests.exceptions.RequestException as e:
        print(f"[ERROR] ADOM {adom_name}: Request error: {e}")
    except (ValueError, KeyError, IndexError, TypeError) as e:
        print(f"[ERROR] ADOM {adom_name}: Invalid API response: {e}")

    return None


def Getdeviceinfo():
    ADOM = ["Retail_ANZ", "Retail_Asia", "Retail_China", "Retail_Japan"]
    #ADOM = ["Retail_ANZ",]
    detaillist = []
    for adom in ADOM:
        info = GetAPI(adom)
        if info is None:
            print(f"[Error] Failed to get device information from {adom} ")
            continue
        # pprint(info)
        for item in info["result"][0]["data"]:
            detaillist.append({
                "adom": adom,
                "hostname": item.get("hostname"),
                "ip": item.get("ip"),
                "platform": item.get("platform_str"),
                "sn": item.get("sn"),
            })
    #pprint(detaillist)
    return detaillist


if __name__ == "__main__":
    start = time.time()
    Getdeviceinfo()
    end = time.time()
    print(f"[TimeCost] Total time: {end - start}")


