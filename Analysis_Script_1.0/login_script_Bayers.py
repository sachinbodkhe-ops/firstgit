import csv
import datetime
import netmiko
import os

def get_logs(host,uname,pass_wd,commands,folder_name):
    cisco_router = {
        'device_type': 'cisco_ios',
        'host': host,
        'fast_cli': False,
        'username': uname,
        'password': pass_wd,
        'secret': ''
    }
    
    try:
        ssh = netmiko.ConnectHandler(**cisco_router)
    except:
        print("connection to "+host+" was NOT successful")
        return False

    print("Connection to "+host+" successful.")

    cmd_output=""
    for command in commands:
        if not command:
            continue
        result = ssh.send_command(command, strip_prompt=False)
        cmd_output=cmd_output+command+"\n"+result
    ssh.disconnect()
    f=open(os.path.join("Logs",folder_name,host+".txt") ,'w')
    cmd=f.write(cmd_output)
    f.close()
    
    f=open(os.path.join("Latest_Log",host+".txt") ,'w')
    cmd=f.write(cmd_output)
    f.close()
    
    print("Logs collected from: "+host+"\n")
    return True

def main():
    try:
        f = open("Commands.txt",'r')
    except:
        print("Commands.txt file is not present.")
        return
    cmd=f.read()
    f.close()
    commands=cmd.split("\n")

    try:
        f = open("Credentials.txt")
        cmd=f.read()
        f.close()
        creds=cmd.split("\n")
        username= creds[0]
        password= creds[1]
    except:
        username=input("Enter the Username: ")
        password=input("Enter the password: ")

    x = datetime.datetime.now()
    folder_name = x.strftime("%Y")+"_"+x.strftime("%m")+"_"+x.strftime("%d")+"_"+x.strftime("%H")+"-"+x.strftime("%M")+"-"+x.strftime("%S")
    if os.path.isdir("Logs"):
        os.mkdir(os.path.join("Logs",folder_name))
    else:
        os.mkdir("Logs")
        os.mkdir(os.path.join("Logs",folder_name))
    
    try:
        f = open("Hosts.txt")
    except:
        print("Hosts.txt file is not present.")
        return
    cmd=f.read()
    f.close()
    hosts=cmd.split("\n")

    for file_ in os.scandir("Latest_Log"):
        os.remove(file_)

    results=[]

    for host in hosts:
        if not host:
            continue
        connection = get_logs(host,username,password,commands,folder_name)
        if connection:
            results.append([host,"Yes"])
        else:
            results.append([host,"No"])
    
    results.insert(0,["Host", "Successful Login"])
    with open("Results.csv", "w", newline="") as f:
        writer = csv.writer(f)
        writer.writerows(results)

if __name__ == "__main__":
    import time
    s = time.perf_counter()
    main()
    print("Total time taken: "+str(time.perf_counter() - s)[:-11])
    print()