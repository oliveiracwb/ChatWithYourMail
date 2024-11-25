# ChatWithYourMail: Export and Interact with Your Outlook Data for Machine Learning 🚀

Welcome to ChatWithYourMail, a powerful Python script designed to transform your Outlook PST data into a rich resource for Machine Learning and RAG (Retriever Augmented Generation) processes. Easily export your emails and attachments to local directories, enabling seamless integration with data processing and analysis workflows.

## 🌟 Features

- **Automated PST Extraction:** Effortlessly extract emails and attachments from Outlook PST files.
- **Structured Directory Organization:** Organizes exported messages and attachments into local directories for easy access.
- **Text Sanitation:** Cleans and structures email text data to ensure quality and consistency.
- **Attachment Handling:** Safely exports all file attachments to designated directories.

## 🔧 Prerequisites

To get started, ensure you have the following:

- Python 3.10.15: The script is built and tested for this version.
- Install the required library:
  ```bash
  pip install pywin32
  ```

## 📄 Usage

Clone this repository to your local machine:
```bash
git clone https://github.com/your-username/chatwithyourmail.git
```

Configure the script:
- Set the `pst` variable to your PST file's path.
- Set the `BASE_DIRECTORY` to where you want the emails and attachments saved.

Run the script:
```bash
python export.py
```

## 🚀 How it Works

The `export.py` script connects with the Outlook application through the pywin32 library. It recursively navigates through your PST file, exporting emails and attachments to the specified local directory while maintaining a structured format.

## Integrating RAG with txtask

To interactively query your exported emails using RAG, follow these steps:

1. **Clone and Prep txtask:**
   ```bash
   git clone https://github.com/your-username/txtask.git
   cd txtask
   ```

2. **Install and Run Ollama:**
   - Ensure you have Ollama installed locally.
   - Run the model, e.g., Mistral 7B:
     ```bash
     ollama run mistral
     ```

3. **Organize Your Exported Emails:**
   - Move or copy the exported email text files to the `./data` folder in txtask.

4. **Run txtask:**
   ```bash
   cargo run
   ```

5. **Start Asking Questions:**
   - Once indexing is complete, you can start querying the email content to gain valuable insights.

## 📂 Directory Structure

The script will save emails and attachments within a specified base directory in the following structure:
```
BASE_DIRECTORY/
├── Folder_Name_1/
│   ├── 1.txt
│   ├── 2.txt
│   └── 1/
│       ├── attachment1.pdf
│       └── attachment2.jpg
├── Folder_Name_2/
│   └── 1.txt
...
```

## 🛠️ Customization

- **Folder Name Sanitization:** Adjust the `sanitize_folder_name()` function for custom directory naming conventions.
- **Text Cleansing:** Customize the `sanitize_text()` function to tweak text formatting and content removal.

## 👥 Contributions

We welcome contributions to enhance the functionality and usability of ChatWithYourMail! If you encounter any issues or have feature requests, please open an issue or submit a pull request.

## 📄 License

This project is licensed under the MIT License. Please refer to the LICENSE file for details.

Enjoy a streamlined transition from Outlook data to Machine Learning insights!
