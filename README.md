# Crawl-Bussiness-Information-from-vinabiz
Example: 

python3 main.py -u https://vinabiz.us/categories/tinhthanh/ha-noi/310030003100 -s 1 -e 2 -o data_clone &

(get data from all bussiness in Ha Noi start from page 1 - 2 and export data_clone.xls file)

Using

Login to vinabiz.us and check network tab by inspect website (or f12)

Tab network

![image](https://user-images.githubusercontent.com/17095935/162409925-00408cc3-d32c-4353-b4d7-add39e87b682.png)

Copy cookie value in header request and replace to cookie value in main.py file

![image](https://user-images.githubusercontent.com/17095935/162410322-9c719019-6f0f-44fb-a5c5-4972d0f33ec9.png)

Save it

Typing cmd in directory bar -> Enter

python main.py -u <link> -s <start_Page> -e <end_Page> -o <name_of_excel_file>

Required Lib

requests

bs4

xlutils

xlwt
