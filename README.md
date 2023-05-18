# System Info Gather 

## 📝 Description
System Info Gather is a script designed to automate the gathering of system information, including hardware and software details, for IT analysis purposes. It provides a convenient way to retrieve various types of system data, saving time and effort for IT analysts. Currently, the script includes network information and system info gathering commands. In addition, there are additional commands mentioned in the "Commands.txt" and "command-book.xlsx" files
<br><br>

## ✨ Features
<lu>
<ul>
  <li>Run on any Windows computer without the need to install additional software or scripts.</li>
  <li>Ability to customize the script to fulfill specific tasks, reducing the need for manual intervention.</li>
  <li>Minimize mouse clicks and copy-paste actions by automating the process of gathering system information.</li>
  <li>Automatically transfer the collected data to a text file for easy organization and analysis.</li>
  <li>Generated text files are named in a proper order, making it convenient to locate and reference specific reports.</li>
</ul>
<br>

<img align=right Width=40% src="https://media1.giphy.com/media/XIahGhbK5A685fyr8D/giphy.gif?cid=ecf05e47htzwy5clysxczuscolhuayx48nq6cehpfr5vnzj0&ep=v1_gifs_search&rid=giphy.gif&ct=g">

## 🛠️ Technology Used
<ul>
  <li>Windows Script Host (WSH)</li>
  <li>VBScript</li>
  <li>CMD commands</li>
  <li>Windows shortcuts</li>
</ul>
<br>

## 📋 Prerequisites
<ul>
  <li>Windows operating system</li>
  <li>Windows Script Host (usually pre-installed on Windows systems)</li>
</ul>
<br>

## 🚀 Installation
<ol>
  <li>Clone or download the System Info Gather repository.</li>
  <li>Run the script file "SystemInfoGather.lnk" using Windows Script Host.</li>
</ol>
<br>

## 💡 Usage
<ol>
  <li>Double-click on the <code>SystemInfoGather.lnk</code> shortcut file.</li>
  <li>The script will open a command prompt window and prompt you to select a directory to store the output files.</li>
  <li>Select the desired directory using the folder browser dialog.</li>
  <li>The script will start gathering system information and executing various commands while displaying.</li>
  <li>Once the process is complete, the script will generate text files with the collected information in the selected directory.</li>
</ol>
<br>

## ⚙️ Configuration
<p>The System Info Gather tool allows you to customize the commands and information you want to gather. Follow these steps to add your own commands:</p>
<ol>
  <li>Open the <code>Tester.vbs</code> file in a text editor.</li>
  <li>Locate the <em>Executable Commands</em> section.</li>
  <li>Add your desired command using the following format:<br>
    <code>PerformCommand "CMD-command-here", "File name"</code></li>
  <li>Save the changes to the <code>Tester.vbs</code> file.</li>
  <li>Run the <code>SystemInfoGather.lnk</code> shortcut file as described in the "Usage" section.</li>
  <li>The modified command will be executed during the script's execution, and the output will be stored in a text file.</li>
</ol>
<p>Repeat the above steps to add more commands or customize the existing ones according to your needs. Ensure that the commands you add are compatible with the Windows Command Prompt environment.</p>
<br>

## 📂 File Structure
<ul>
  <li><code>SystemInfoGather.lnk</code></li>
  <li><code>Tester.vbs</code></li>
  <li><code>Commands.txt</code></li>
  <li><code>Command-book.xlsx</code></li>
  <li><code>README.md</code></li>
</ul>
<br>

## 🚦 Troubleshooting
<li>If you encounter any errors related to insufficient privileges or access denied, it may be due to the script requiring administrative privileges to gather certain system information. In such cases, you can try running the script with administrative rights to resolve the issue.</li>
<li>To run the script as an administrator:</li>
<ol>
  <li>Right-click on the <code>SystemInfoGather.lnk</code> shortcut file.</li>
  <li>Select "Run as administrator" from the context menu.</li>
</ol>
<li>By running the script with administrative privileges, you ensure that it has the necessary permissions to access and retrieve the desired system information.</li>

<br>

## 🤝 Contributions
Contributions to the System Info Gather project are welcome. Please follow the guidelines outlined in the CONTRIBUTING.md file.
<br><br>

## 📄 License
This project is licensed under the MIT License.
<br><br>

## 🙏 Acknowledgements
Special thanks to the contributors and the open-source community for their support and valuable feedback.