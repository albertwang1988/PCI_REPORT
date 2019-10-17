# PCI-REPORT Manual



**PCI-REPORT** is  a application to manage records uploaded to *mta.org* of patients undergoing percutaneous coronary intervention. This project is based on Python 3 and is operational in all platform(Windows, MacOS and Linux,etc.). Addtional component may be needed via *pip3*.

Python 3 virtual enviroment is **recommended**, to void confilict with pip package already installed in your server.

PCI-REPORT uses xlsx file to manage all data. The xlsx file can be opened by *Microsoft Excel* or *IBM SPSS*. Make sure to **backup your data at least once a month**.

---

### Function

- Add a record

  use `pcir add` command to add an record to the database. This command will create a new record with a query of patient's ID, name and status. The record will be written in the database in the format of : sn, ID, name, gender, status.

- Show unsubmitted records

  use `pcir show` to show all the records that have not been submitted to *mta.org*.

- Show all valid records

  use `pcir showall` to show all the records within the database. 

  **MENTION:** record with a 'DEL' mark will not be shown.

- Show all records

  use `pcir showfull` to show all the records including those with a 'DEL' mark.

- Change status

  use `pcir change [sn] ` to change the status of submition of any record. the status would only change from 'Y' to 'N', or otherwise.

- Delete a record

  use `pcir del [sn]`  command to delete an record to the database. This command will ask for the serial number of the record. This command only marked the record as 'DEL', not really erase the record from the xlsx file.
  
- Undelete a record

  use `pcir undnel [sn]` command to cancel the 'DEL' mark of a record. The mark will go back to 'N' by defualt.

- Erase all record marked as 'DEL'

  use `pcir erase` to erase all records marked as 'DEL'. Records will be missing instantly. Please double check before you apply this command.

- Show help informatiln

  Simply use `pcir` or `pcir -help` to show help information. You will see this document.



### Argument

- add
- show
- showall
- showfull
- change
- del
- undel
- erase
- -help