### **Change Partition**
To use 'cd' to change directory in another partetion, you should use 'cd d/'
```bash
pwd
E:\Programs

cd /d D:\FolderDiscovery
D:\FolderDiscovery
```

### **How to switch Python versions in Windows 10. Set Python path**

Link: https://www.youtube.com/watch?v=zriWqGNJg4k
```bash
py -0                   # Show versions of installed python
Installed Pythons found by py Launcher for Windows
 -3.10-64 *
 -3.9-64

python --version        # Get python used version
Python 3.9.7

py -3.10-64             # run python version 3.10-64
Python 3.10.11 (tags/v3.10.11:7d4cc5a, Apr  5 2023, 00:38:17) [MSC v.1929 64 bit (AMD64)] on win32
Type "help", "copyright", "credits" or "license" for more information.
>>> exit()              # Exit Python not terminal
```
1. Right click on windows at toolbar button.
2. Select **System**.
3. From **Related settings** select **Advanced system settings**.
4. Select **Enviromment Variables**
5. Select **Path** and click **Edit**<br>
   ***Hint:*** *Edit Path in **System variable** not **User variable**.*<br>
   ***Hint Link:*** https://www.c-sharpcorner.com/article/how-to-manage-multiple-versions-of-python-on-windows-11/
6. Select python path and up to the top one.

### **Setup Virtual Environment**
- Link: https://www.geeksforgeeks.org/creating-python-virtual-environment-windows-linux/
- Link: https://www.youtube.com/watch?v=oL6pwOK253I
```bash
cd /d C:\Users\<PC User Name>
pip install virtualenv        # Install vertual environment
virtualenv foldis             # Create 'foldis' virtual environment
python -m virtualenv foldis   # Create 'foldis' virtual environment (Another Method)
foldis\Scripts\activate       # Activate 'foldis' env
(foldis)                      # Now you activate env 'foldis' and inside it
pip freeze                    # Show libraries that is installed in env
deactivate                    # Deactivate env
```
### **Setup PyQt5**
- Link: https://www.techwithtim.net/tutorials/pyqt5-tutorial/how-to-use-qtdesigner/
- Link: https://gist.github.com/marcoandre1/a77460d7b88de7e9608335b9c518b752
```bash
cd /d C:\Users\<PC User Name>
foldis\Scripts\activate
pip3 install pyqt5            # Install PyQt5
pip freeze
PyQt5==5.15.9
PyQt5-Qt5==5.15.2
PyQt5-sip==12.12.1

pip install pyqt5-tools       # Install PyQt5 Designer 'designer.exe'
pip freeze
click==8.1.3
colorama==0.4.6
PyQt5==5.15.9
pyqt5-plugins==5.15.9.2.3
PyQt5-Qt5==5.15.2
PyQt5-sip==12.12.1
pyqt5-tools==5.15.9.3.3
python-dotenv==1.0.0
qt5-applications==5.15.2.2.3
qt5-tools==5.15.2.1.3

designer.exe                  # Run PyQt5 Designer, it's located in: C:\Users\<PC User Name>\anaconda3\Library\bin\designer.exe
```

### **Setup Other Resources**
```bash
pip install pandas            # Install Pandas
pip freeze
click==8.1.3
colorama==0.4.6
numpy==1.24.3
pandas==2.0.1
PyQt5==5.15.9
pyqt5-plugins==5.15.9.2.3
PyQt5-Qt5==5.15.2
PyQt5-sip==12.12.1
pyqt5-tools==5.15.9.3.3
python-dateutil==2.8.2
python-dotenv==1.0.0
pytz==2023.3
qt5-applications==5.15.2.2.3
qt5-tools==5.15.2.1.3
six==1.16.0
tzdata==2023.3

pip install openpyxl          # Install OpenPyXl
pip freeze
click==8.1.3
colorama==0.4.6
et-xmlfile==1.1.0
numpy==1.24.3
openpyxl==3.1.2
pandas==2.0.1
PyQt5==5.15.9
pyqt5-plugins==5.15.9.2.3
PyQt5-Qt5==5.15.2
PyQt5-sip==12.12.1
pyqt5-tools==5.15.9.3.3
python-dateutil==2.8.2
python-dotenv==1.0.0
pytz==2023.3
qt5-applications==5.15.2.2.3
qt5-tools==5.15.2.1.3
six==1.16.0
tzdata==2023.3
```

### **'ctypes' Module**
- Link: https://stackoverflow.com/questions/14900510/changing-the-application-and-taskbar-icon-python-tkinter
- Link: https://docs.python.org/3/library/ctypes.html

A foreign function library for Python. ctypes is a foreign function library for Python. It provides C compatible data types, and allows calling functions in DLLs or shared libraries. It can be used to wrap these libraries in pure Python.

#### **Changing the application and taskbar icon - Python/Tkinter/PyQt**
How can I change its application icon (the 'file' icon shown at the explorer window and the start/all programs window, for example - not the 'file type' icon nor the main window of the app icon) and the taskbar icon (the icon shown at the taskbar when the application is minimized)? I only need to support Windows XP and Win7 machines.
#### **Answer**
Another option on Windows would be the following:

To your python code add the following:
```python
import ctypes

myappid = 'mycompany.myproduct.subproduct.version' # arbitrary string
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
```

### **Run Code**
```python
python FolDis-MW_V0.10.py
```

### **Setup PyInstaller**
```bash
pip install pyinstaller       # Install PyInstaller
pip freeze
altgraph==0.17.3
click==8.1.3
colorama==0.4.6
et-xmlfile==1.1.0
numpy==1.24.3
openpyxl==3.1.2
pandas==2.0.1
pefile==2023.2.7
pyinstaller==5.10.1
pyinstaller-hooks-contrib==2023.2
PyQt5==5.15.9
pyqt5-plugins==5.15.9.2.3
PyQt5-Qt5==5.15.2
PyQt5-sip==12.12.1
pyqt5-tools==5.15.9.3.3
python-dateutil==2.8.2
python-dotenv==1.0.0
pytz==2023.3
pywin32-ctypes==0.2.0
qt5-applications==5.15.2.2.3
qt5-tools==5.15.2.1.3
six==1.16.0
tzdata==2023.3

pyinstaller --windowed -F FolDis-MW_V0.10.py    # Convert 'FolDis-MW_V0.10.py' to 'exe' file, '-F' to create a one-file bundled executable.
# Hint: The 'exe' file is located in folder 'dist', and you should copy image/icon files to this folder.
```