import uno
import json
import datetime
import paho.mqtt.client as paho

def detect_last_used_row(sheet_name):
    global doc
    oSheet = doc.Sheets.getByName(sheet_name)
    oCursor = oSheet.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    return oCursor.getRangeAddress().EndRow +1

def detect_last_used_column(sheet_name):
    global doc
    oSheet = doc.Sheets.getByName(sheet_name)
    oCursor = oSheet.createCursor()
    oCursor.gotoEndOfUsedArea(False)
    return oCursor.getRangeAddress().EndColumn +1

def on_connect(client, userdata, flags, rc):
    log_message("rc: " + str(rc))

def on_subscribe(client, obj, mid, granted_qos):
    log_message("Subscribed: " + str(mid) + " " + str(granted_qos))

def on_log(client, obj, level, string):
    log_message(string)

def log_message(message):
    datetime.datetime.now
    if doc.Sheets.hasByName("logs"):
        sheet = doc.Sheets.getByName("logs")
        last_row = detect_last_used_row("logs")
        sheet.getCellByPosition(0, last_row).setString(str(datetime.datetime.now()))
        sheet.getCellByPosition(1, last_row).setString(message)
    else:
        print(str(datetime.datetime.now()) + " | " + message)

def fill_data(client, userdata, message):
    global doc
    dict_data=json.loads(message.payload.decode("utf-8"))
    sheet = doc.Sheets.getByName("data")
    last_row = detect_last_used_row("data")
    last_column = detect_last_used_column("data")
    for i in range(0, last_column):
        index=sheet.getCellByPosition(i, 0).getString()
        if index.startswith("#") == False:
            if index in dict_data.keys():
                sheet.getCellByPosition(i, last_row).setString(dict_data[index])
            else:
                sheet.getCellByPosition(i, last_row).setString("null")

def launch_job():
    try:
        global doc
        client= paho.Client()
        client.on_connect = on_connect
        client.on_subscribe = on_subscribe
        client.on_log = on_log
        client.on_message=fill_data
        sheet=doc.Sheets.getByName("parameters")
        broker=sheet.getCellByPosition(1, 1).getString()
        topic=sheet.getCellByPosition(1, 2).getString()
        login=sheet.getCellByPosition(1, 3).getString()
        password=sheet.getCellByPosition(1, 4).getString()
        client.username_pw_set(login, password)
        client.connect(broker)
        client.subscribe(topic)
        client.loop_forever()
    except KeyboardInterrupt:
        client.loop_stop()

localContext = uno.getComponentContext()
resolver = localContext.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", localContext )
ctx = resolver.resolve( "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext" )
smgr = ctx.ServiceManager
desktop = smgr.createInstanceWithContext( "com.sun.star.frame.Desktop",ctx)
doc = desktop.getCurrentComponent()
launch_job()
ctx.ServiceManager
