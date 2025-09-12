# 📊 Automated Monthly Task & Jira Report Generator

![Java](https://img.shields.io/badge/Java-17%2B-blue)
![Maven](https://img.shields.io/badge/Maven-Build-orange)
![ApachePOI](https://img.shields.io/badge/ApachePOI-Excel-green)
![License](https://img.shields.io/badge/License-MIT-yellow)
![Status](https://img.shields.io/badge/Status-Stable-success)

---

## 🚀 Overview

This project is a **Java-based automation tool** that reduces manual reporting efforts by up to **80%**.
It reads daily task `.txt` files, extracts task details with **Regex**, and generates a **monthly Excel report** using **Apache POI**.

Packaged as a **reusable JAR**, it runs anywhere with Java installed — just drop it in your folder with `.txt` files and run.

---

## ✨ Features

* 📅 Extracts **date** from `.txt` file names
* 📝 Extracts **tasks + Jira IDs** from file contents
* 📊 Generates a **monthly Excel report** with:

  * Name
  * Manager Name
  * Employee ID
  * Project Name
  * Month & Year
* ⚡ Saves up to **80% reporting time**
* 🔁 Packaged as a **reusable JAR**

---

## 🛠️ Tech Stack

* **Java 17+**
* **Apache POI** (Excel handling)
* **Regex** (task extraction)
* **File I/O**
* **Maven** (build & dependencies)

---

## 📂 Project Structure

```
project-root/
│── input.txt        # Contains user details (Name, Manager, Employee ID, Project, Month, Year)  
│── 2025-09-01.txt   # Example daily log file (file name = date, contents = tasks & Jira IDs)  
│── 2025-09-02.txt  
│── target/          # Maven build output  
│── report-generator.jar  # Generated reusable JAR  
```

---

## 📥 Installation & Usage

1️⃣ **Clone the repo:**

```bash
git clone https://github.com/your-username/task-jira-report-generator.git
cd task-jira-report-generator
```

2️⃣ **Build the project (Maven):**

```bash
mvn clean package
```

3️⃣ **Prepare input files:**

* Create an `input.txt` with details like:

  ```
  Name=John Doe
  Manager=Jane Smith
  EmployeeID=EMP123
  Project=ProjectX
  Month=September
  Year=2025
  ```
* Add your daily log `.txt` files (e.g., `2025-09-01.txt`) with tasks + Jira IDs inside.

4️⃣ **Run the JAR:**

```bash
java -jar target/report-generator.jar
```

✅ Your monthly Excel report will be generated automatically in the same folder!

---

## 📸 Example

**Input (`2025-09-01.txt`):**

```
TASK-101 Fix login bug
TASK-102 Implement dashboard
```

**Output (Excel):**

| Date       | Task                | Jira ID  |
| ---------- | ------------------- | -------- |
| 2025-09-01 | Fix login bug       | TASK-101 |
| 2025-09-01 | Implement dashboard | TASK-102 |

---

## 🤝 Contributing

Contributions, issues, and feature requests are welcome!
Feel free to fork this repo and submit a PR.

---

## 📜 License

This project is licensed under the **MIT License**.

---

👉 Do you also want me to make a **demo GIF / screenshot section** (showing the folder + generated Excel) so visitors instantly get what your project does?
