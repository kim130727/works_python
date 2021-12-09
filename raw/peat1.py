#!/usr/bin/env python
# -*- coding: utf-8 -*-
#

import os
import requests

ip = "api.peat-cloud.com"
version = "v1"
route = "image_analysis"
url = "http://%s/%s/%s" % (ip, version, route)


def single_processing():
    # Header of our requst. Replace <YOUR_API_KEY> with your api key.
    headers = {"api_key": GUEST, "variety": "TOMATO"}

    # make a dict with the picture
    image = "data/tomato_nutrient/iron1.jpg"
    files = {"picture": open(image, "rb")}

    # post both files to our API
    result = requests.get(url, files=files, headers=headers, timeout=10)

    if result.status_code == 401:
        print ("Authentication failed")
    elif result.status_code == 500:
        print ("Internal server error...")
    elif result.status_code == 200:
        # load response that comes in JSON format and print the result
        json_data = result.json()
        for data in json_data["image_analysis"]:
            print ("Disease name: %s\n\tProbability: %s%%" % (data["name"], data["similarity"]))
    print ("")
    return

if __name__ == "__main__":
    single_processing()