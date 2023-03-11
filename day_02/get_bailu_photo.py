"""
@author:jsp
@file:get_bailu_photo.py
@time:2023/3/11 18:22
@desc:
"""
import requests

photo=requests.get("https://lmg.jj20.com/up/allimg/mx10/010120234309/200101234309-0-lp.jpg")
print(photo.content)
file=open("白鹿.jpg","wb")
file.write(photo.content)
file.close()