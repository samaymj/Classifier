# Features (Version 1 – Rule-Based Classifier)

###  **1. Rule-Based Ticket Classification**

Automatically assigns each support ticket to the correct category by matching keywords from a predefined keyword list.

###  **2. Text Cleaning & Normalization**

Preprocessing applied to improve matching accuracy:
* Lowercasing
* Removing punctuation
* Removing extra spaces
* Normalized text stored as a separate column

###  **3. Keyword Mapping from Excel**

Reads `category_keywords.xlsx` and builds a category → keywords dictionary dynamically.
Supports multiple categories and unlimited keywords.

###  **4. Handles Unclassified Tickets**

Any ticket that does not match any keyword is automatically assigned to: 'others'
A separate list of all unclassified ticket IDs and descriptions is included in the summary.

###  **5. Multiple Keyword Matches (Configurable)**

* You can choose to assign the **first matched category**, OR
* Store **multiple categories** (configurable through a single variable).

###  **6. Clean & Structured Outputs**

The script generates:
#### **A) tickets_classified.xlsx**
Contains:
    * ticket_id
    * original description
    * cleaned description
    * assigned category
    * matched categories list
Easy for further analysis or reporting.

#### **B) summary.json**

Contains a detailed summary:
    
    * Total tickets processed
    * Ticket count per category
    * Number of unclassified tickets
    * Full list of unclassified ticket details

###  **7. Works on Any Excel Dataset**

As long as the files contain:

* `ticket_id`, `description` (in tickets.xlsx)
* `category`, `keywords` (in category_keywords.xlsx)
