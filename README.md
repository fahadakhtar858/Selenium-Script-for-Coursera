
title: Selenium Script for Coursera and EDx

---

# Selenium Script for Coursera

The project will scrap the Coursera to gather the dataset containing:

-   Course Name
-   Course URL
-   Course Instructors
-   Designation of the Course Instructor
-   Course Offering Institute
-   Department of the Institute
-   Skills Offered in the Course
-   Course Rating
-   Course Reviews
-   Number of Enrolments in the Course
-   Course Level
-   Course Type
-   Course Duration

The script is designed to scrape the first five pages of the search results and compile them into an `xlsx` file based on a search string.

Below are the steps to run the script and gather results:

Step 1: We’ll need the following two software to run the script:

    - Java 1.8
    - Chrome Driver

        Step 1: We need Java 1.8 on our machine. We'll use this [<u>link</u>](https://java.com/en/download/) to download and install Java.

        Step 2: Following are the steps to install Chrome Web Driver:

    - Type the following command in Google Chrome's address bar to get its installed version on your machine:

                `chrome://settings/help`

            Now, we need to download the Chrome Driver which is compatible with our installed Chrome browser. Use this [<u>link</u>](https://chromedriver.chromium.org/downloads) to download the version of the Chrome Driver which matches your Chrome browser's version.

    - Extract the downloaded chrome driver and place it in the `Path`  folder (`/usr/local/bin/`). In order to do so, open the terminal where the chrome driver is extracted and type the following command:

                `sudo cp chromedriver /usr/local/bin/ `

    - Navigate the terminal to the `/usr/local/bin/`  location and run the following command:

                `sudo xattr -d com.apple.quarantine chromedriver`

Step 2: Following are the steps to run the executable jar:

    - Open the terminal in the same location as your jar file and type the following commands to execute it: java -jar Coursera.jar
    - Once running it will ask for the search string and path to save the excel file.
    - Enter the search string e.g. cyber security, deep learning, blockchain, etc.
    - Enter the path in the following format: `/Users/Fahad//Desktop/`

        Here, `Fahad`  is the username of the machine. Update it with the username of the machine where this script is being executed. At the end of the run, it will create a `.xlsx`  file at your provided path.

**Important platform-specific attributes:**

-   Coursera 12 search results/page
-   EDx 24 search results/page

          
