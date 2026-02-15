import json

credentials = {
    "cox_url": "https://cox.fs.ocs.oraclecloud.com/",
    "username": "c50105",
    "password": "Coxdispatch2026!",
    "Mail": "stephen.davis@cox.com",
    "Mail_password": "Coxdispatch2026!",
    "FuseURL": "https://fuse.i-t-g.net/login.php",
    "Fusername": "ybot",
    "Fpassword": "Bluebird1@3",
    "Logfile": r"C:\Users\RohansarathyGoudhama\Downloads\cox\cox_work_orders_logs", 
    "ybotID":"ybot@itgext.com",
    "sdavis":"sdavis@itgcomm.com",
    "Default_Path":r"C:\Users\RohansarathyGoudhama\Downloads\cox",
}

with open("cox_work_order.json", "w") as json_file:
    json.dump(credentials, json_file)