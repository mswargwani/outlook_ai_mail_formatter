# OpenRouter AI for Outlook

A Microsoft Outlook VBA macro that integrates OpenRouter's free AI model (`openrouter/free`) directly into your email client. This tool acts as a personal AI assistant, helping you draft new emails, generate context-aware replies, and refine your text with just a click.

## Features

* **Zero Cost:** Powered by OpenRouter's free AI models.
* **Smart Replies:** Generate context-aware replies with a single click (Yes, No, Thanks, Follow-up, Acknowledge).
* **Compose New Mail:** Quickly draft professional emails based on a short topic prompt.
* **Text Processing Tools:** Highlight text within any email to reword it, explain it, or apply a custom AI prompt.
* **Seamless Integration:** Automatically picks up your Outlook username and default email signature.
* **Secure API Key Storage:** Prompts for your API key only on the first run and saves it securely in your local registry.

---

## Prerequisites

1. **Microsoft Outlook** (Desktop Client for Windows)
2. **OpenRouter API Key:** You can get a free API key by signing up at [https://openrouter.ai/keys](https://openrouter.ai/keys)

---

## Installation

1. Open **Microsoft Outlook**.
2. Press `Alt + F11` to open the VBA Editor.
3. In the top menu, go to **Insert > Module**.
4. Paste the entire VBA script into the new module window.
5. Close the VBA Editor.

### Adding Macros to your Ribbon (Recommended)

To make the tools easily accessible, add them to your Outlook ribbon:

1. Right-click the Outlook Ribbon and select **Customize the Ribbon**.
2. Create a **New Group** under the *Home (Mail)* tab (name it "AI Tools").
3. Under the *Choose commands from* dropdown, select **Macros**.
4. Drag and drop the macros (e.g., `Project1.ComposeNewMail`, `Project1.ReplyYes`) into your new custom group.
5. Rename them and assign icons for a cleaner look.

---

## Usage Guide

The first time you run any of these macros, you will be prompted to paste your **OpenRouter API Key**. It will be saved automatically for future use.

### 1. Generating Replies

Select an email in your inbox and run one of the following macros. The AI will read the original email and draft a response for you, appending your default signature at the bottom.

* **`ReplyYes`**: Writes an affirmative/positive response.
* **`ReplyNo`**: Writes a polite declining response.
* **`ReplyThanks`**: Writes a thank-you note.
* **`ReplyFollowup`**: Writes a follow-up response.
* **`AcknowledgeEmail`**: Lets the sender know you received their message and will reply later.

### 2. Composing New Mail

* **`ComposeNewMail`**: Prompts you for a short description (e.g., "Ask the marketing team for the Q3 report"). The AI will generate a complete professional email, including a subject line, and open it in a new window.

### 3. Text Processing

Open an email, highlight specific text, and run:

* **`Reword`**: Improves the grammar, clarity, and professionalism of the selected text.
* **`Explain`**: Breaks down complex highlighted text into simple terms.
* **`GPTText`**: Opens a prompt box asking what you want the AI to do with the highlighted text (e.g., "Translate to French", "Summarize into 3 bullet points").

---

## Configuration

You can customize the AI's behavior by modifying the `Const` variables at the top of the VBA script:

| Constant | Default Value | Description |
| --- | --- | --- |
| `MODEL_NAME` | `"openrouter/free"` | The AI model used. Can be changed to premium models if you have credit. |
| `TONE` | `"friendly"` | The tone of the generated emails (e.g., "professional", "casual", "urgent"). |
| `LANGUAGE` | `"English"` | The default language for all outputs. |
| `MAX_TOKENS` | `"2048"` | Maximum length of the AI's response. |

---

## Troubleshooting

* **TestConnection:** Run this macro to verify that your OpenRouter API key is valid and that Outlook can successfully communicate with the OpenRouter servers.
* **FactoryReset:** If you need to change your API key or accidentally pasted the wrong one, run this macro. It will wipe the saved key from your registry and prompt you for a new one the next time you run a command.
* **Network/Library Errors:** Ensure your IT department does not block `MSXML2.XMLHTTP` or `WinHttp.WinHttpRequest` traffic in Microsoft Office applications.
