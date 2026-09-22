/************************************************************
 * QA AUTOMATION RECRUITER EMAIL SYSTEM
 * Harshad Chandrakant Shinde
 *
 * ONLY 4 RUNNABLE FUNCTIONS SHOULD APPEAR IN THE
 * APPS SCRIPT FUNCTION DROPDOWN:
 *
 * 1. sendRecruiterEmails
 * 2. followUpDay02
 * 3. followUpDay03
 * 4. followUpDay04
 *
 * NEW FEATURES
 * ----------------------------------------------------------
 * - Tracks Gmail delivery failures / bounces / blocks
 * - Marks bounced/blocked recruiter emails RED
 * - Marks successfully sent/no-bounce emails GREEN
 * - Tracks recruiter replies across Day 01-Day 04
 * - Marks replied recruiter emails BLUE
 * - Stops future follow-ups after a recruiter replies
 * - Stops future follow-ups after bounce/block
 * - Moves matched Gmail delivery-failure notifications to Trash
 * - Tracks last recruiter response date
 * - Keeps Gmail follow-ups in the same conversation thread
 * - Uses Gmail API for explicit To + threadId + reply headers
 * - Displays clear Gmail quota information in the log
 *
 * IMPORTANT:
 * "Delivered" means "sent successfully and no Gmail delivery
 * failure/bounce has been detected by the next status scan".
 * Gmail does not expose a simple SMTP delivery-confirmation
 * event to Apps Script. A later bounce can change GREEN to RED.
 ************************************************************/


/************************************************************
 * CONFIGURATION
 ************************************************************/

const CONFIG = {

  MAX_INITIAL_EMAILS_PER_RUN: 10,

  MAX_FOLLOWUPS_PER_RUN: 10,

  RESUME_FILE_NAME:
    "HARSHAD SHINDE_RESUME.pdf",

  EMAIL_SUBJECT:
    "QA Automation Engineer | 5.7 Years | Harshad Shinde",

  SENDER_NAME:
    "Harshad Shinde",

  TEST_COMPANY_NAME:
    "TEST",

  TIMEZONE:
    "Asia/Kolkata",

  BOUNCE_LOOKBACK_DAYS: 30,

  MAX_BOUNCE_THREADS_PER_QUERY: 100,

  COLORS: {
    GREEN: "#d9ead3",
    RED: "#f4cccc",
    BLUE: "#cfe2f3",
    YELLOW: "#fff2cc",
    WHITE: "#ffffff"
  },

  // Leave blank unless automatic sender-email detection fails.
  SENDER_EMAIL: ""
};


/************************************************************
 * PUBLIC FUNCTION 1
 * SEND INITIAL RECRUITER EMAILS
 ************************************************************/

function sendRecruiterEmails() {

  if (!PRIVATE.isBusinessDay()) {

    Logger.log(
      "Today is Saturday/Sunday. No initial emails will be sent."
    );

    return;
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getActiveSheet();

  PRIVATE.ensureColumns(sheet);

  // Detect existing replies and Gmail delivery failures first.
  const syncSummary =
    PRIVATE.syncEmailStatuses(sheet);

  const data =
    sheet
      .getDataRange()
      .getValues();

  if (data.length < 2) {

    Logger.log(
      "No recruiter rows found."
    );

    return;
  }

  const headers = data[0];

  const companyCol =
    PRIVATE.findColumn(headers, ["Company"]);

  const emailCol =
    PRIVATE.findColumn(headers, ["Email"]);

  const recruiterCol =
    PRIVATE.findColumn(
      headers,
      ["Recruiter Name", "Recruiter"]
    );

  const statusCol =
    PRIVATE.findColumn(headers, ["Status"]);

  const sentDateCol =
    PRIVATE.findColumn(headers, ["Sent Date"]);

  const threadIdCol =
    PRIVATE.findColumn(headers, ["Thread ID"]);

  const deliveryCol =
    PRIVATE.findColumn(headers, ["Delivery Status"]);

  const responseCol =
    PRIVATE.findColumn(headers, ["Response Status"]);

  if (
    companyCol === -1 ||
    emailCol === -1 ||
    recruiterCol === -1 ||
    statusCol === -1 ||
    sentDateCol === -1 ||
    threadIdCol === -1 ||
    deliveryCol === -1 ||
    responseCol === -1
  ) {

    throw new Error(
      "Required columns are missing."
    );
  }


  /**********************************************************
   * FIND RESUME
   **********************************************************/

  const resumeFiles =
    DriveApp
      .getFilesByName(
        CONFIG.RESUME_FILE_NAME
      );

  if (!resumeFiles.hasNext()) {

    throw new Error(
      "Resume not found in Google Drive: " +
      CONFIG.RESUME_FILE_NAME
    );
  }

  const resumeFile =
    resumeFiles.next();


  /**********************************************************
   * CHECK DAILY QUOTA
   **********************************************************/

  let remainingQuota =
    MailApp.getRemainingDailyQuota();

  Logger.log(
    "========================================"
  );

  Logger.log(
    "EMAIL AUTOMATION STATUS"
  );

  Logger.log(
    "========================================"
  );

  Logger.log(
    "Gmail recipient quota remaining: " +
    remainingQuota
  );

  Logger.log(
    "Bounce/blocked notifications detected: " +
    syncSummary.failures
  );

  Logger.log(
    "Recruiter replies detected: " +
    syncSummary.replies
  );

  Logger.log(
    "========================================"
  );


  if (remainingQuota <= 0) {

    Logger.log(
      "No Gmail recipient quota remaining today."
    );

    return;
  }

  let emailsSent = 0;


  /**********************************************************
   * PROCESS RECRUITERS
   **********************************************************/

  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      emailsSent >=
      CONFIG.MAX_INITIAL_EMAILS_PER_RUN
    ) {
      break;
    }

    if (
      remainingQuota <= 0
    ) {
      break;
    }

    const rowNumber = i + 1;

    const company =
      String(
        data[i][companyCol] || ""
      ).trim();

    const email =
      PRIVATE.normalizeEmail(
        data[i][emailCol]
      );

    let recruiterName =
      String(
        data[i][recruiterCol] || ""
      ).trim();

    const status =
      String(
        data[i][statusCol] || ""
      )
        .trim()
        .toLowerCase();

    const existingThreadId =
      String(
        data[i][threadIdCol] || ""
      ).trim();

    const deliveryStatus =
      String(
        data[i][deliveryCol] || ""
      )
        .trim()
        .toLowerCase();

    const responseStatus =
      String(
        data[i][responseCol] || ""
      )
        .trim()
        .toLowerCase();


    /********************************************************
     * VALIDATION
     ********************************************************/

    if (!email) {
      continue;
    }


    if (
      company.toUpperCase() ===
      CONFIG.TEST_COMPANY_NAME
    ) {

      Logger.log(
        "Skipping TEST row: " + email
      );

      continue;
    }


    /********************************************************
     * NEVER SEND TO BOUNCED/BLOCKED/REPLIED CONTACT
     ********************************************************/

    if (
      deliveryStatus === "bounced" ||
      deliveryStatus === "blocked" ||
      responseStatus === "replied"
    ) {

      continue;
    }


    /********************************************************
     * SENT = NEVER SEND INITIAL EMAIL AGAIN
     ********************************************************/

    if (
      status === "sent"
    ) {

      continue;
    }


    /********************************************************
     * ELIGIBLE INITIAL STATUSES
     ********************************************************/

    if (
      status !== "" &&
      status !== "pending" &&
      status !== "failed"
    ) {

      continue;
    }


    /********************************************************
     * IF THREAD ID ALREADY EXISTS,
     * DO NOT SEND INITIAL EMAIL AGAIN.
     ********************************************************/

    if (
      existingThreadId &&
      existingThreadId !== "PENDING_RECOVERY"
    ) {

      Logger.log(
        "Skipping already processed row: " + email
      );

      continue;
    }


    /********************************************************
     * RECRUITER NAME
     ********************************************************/

    if (!recruiterName) {

      recruiterName =
        PRIVATE.deriveRecruiterName(email);

      if (
        recruiterName !== "Team"
      ) {

        sheet
          .getRange(
            rowNumber,
            recruiterCol + 1
          )
          .setValue(
            recruiterName
          );
      }
    }


    /********************************************************
     * INITIAL EMAIL BODY
     ********************************************************/

    const emailBody =
`Hi ${recruiterName},

I hope you’re doing well.

My name is Harshad Chandrakant Shinde, and I am currently looking for opportunities in QA Automation Testing.

I have 5.7 years of experience in software testing, including 3.7 years in automation testing using Java, Selenium WebDriver, TestNG, Cucumber BDD, Playwright, and API testing with Postman & Rest Assured. I also have 2 years of manual testing experience.

I worked at Xenneo Tech Private Limited, where I worked on automation framework development, regression testing, API validation, and Agile-based projects.

Please find my details below:

Last Company: Xenneo Tech Private Limited
Current Location: Pune
Notice Period: Already Served [27 Feb 2026].

Please consider my profile for any suitable QA Automation Engineer openings.

Please find my updated resume attached for your reference.

Regards,
Harshad Chandrakant Shinde
QA Automation Engineer
Pune
+91-7620215494
harshadshinde.work@gmail.com`;


    /********************************************************
     * SEND EMAIL
     ********************************************************/

    try {

      const draft =
        GmailApp.createDraft(
          email,
          CONFIG.EMAIL_SUBJECT,
          emailBody,
          {
            attachments: [
              resumeFile.getBlob()
            ],
            name: CONFIG.SENDER_NAME
          }
        );

      const sentMessage =
        draft.send();

      const thread =
        sentMessage.getThread();

      if (!thread) {

        throw new Error(
          "Email was sent but Gmail Thread ID could not be captured."
        );
      }

      const threadId =
        thread.getId();

      const sendDate =
        new Date();


      /********************************************************
       * UPDATE INITIAL STATUS
       ********************************************************/

      sheet
        .getRange(
          rowNumber,
          statusCol + 1
        )
        .setValue("Sent");


      sheet
        .getRange(
          rowNumber,
          sentDateCol + 1
        )
        .setValue(sendDate);


      sheet
        .getRange(
          rowNumber,
          threadIdCol + 1
        )
        .setValue(threadId);


      // Green means Gmail accepted the message and no
      // delivery failure has been detected yet.
      sheet
        .getRange(
          rowNumber,
          deliveryCol + 1
        )
        .setValue("Delivered");


      sheet
        .getRange(
          rowNumber,
          responseCol + 1
        )
        .setValue("Waiting");


      /********************************************************
       * FOLLOW-UP COLUMNS
       ********************************************************/

      const f2StatusCol =
        PRIVATE.findColumn(
          headers,
          ["Follow-up 02 Status"]
        );

      const f2DateCol =
        PRIVATE.findColumn(
          headers,
          ["Follow-up 02 Date"]
        );

      const f3StatusCol =
        PRIVATE.findColumn(
          headers,
          ["Follow-up 03 Status"]
        );

      const f3DateCol =
        PRIVATE.findColumn(
          headers,
          ["Follow-up 03 Date"]
        );

      const f4StatusCol =
        PRIVATE.findColumn(
          headers,
          ["Follow-up 04 Status"]
        );

      const f4DateCol =
        PRIVATE.findColumn(
          headers,
          ["Follow-up 04 Date"]
        );


      const day02 =
        PRIVATE.addBusinessDays(
          sendDate,
          1
        );

      const day03 =
        PRIVATE.addBusinessDays(
          sendDate,
          2
        );

      const day04 =
        PRIVATE.addBusinessDays(
          sendDate,
          3
        );


      sheet
        .getRange(
          rowNumber,
          f2StatusCol + 1
        )
        .setValue("Pending");

      sheet
        .getRange(
          rowNumber,
          f2DateCol + 1
        )
        .setValue(day02);


      sheet
        .getRange(
          rowNumber,
          f3StatusCol + 1
        )
        .setValue("Pending");

      sheet
        .getRange(
          rowNumber,
          f3DateCol + 1
        )
        .setValue(day03);


      sheet
        .getRange(
          rowNumber,
          f4StatusCol + 1
        )
        .setValue("Pending");

      sheet
        .getRange(
          rowNumber,
          f4DateCol + 1
        )
        .setValue(day04);


      PRIVATE.applyRowColor(
        sheet,
        rowNumber,
        headers
      );


      emailsSent++;
      remainingQuota--;

      Logger.log(
        "Initial email sent successfully: " +
        email +
        " | Recruiter: " +
        recruiterName +
        " | Thread ID: " +
        threadId
      );

    } catch (error) {

      sheet
        .getRange(
          rowNumber,
          statusCol + 1
        )
        .setValue("Failed");

      sheet
        .getRange(
          rowNumber,
          deliveryCol + 1
        )
        .setValue("Failed");

      sheet
        .getRange(
          rowNumber,
          responseCol + 1
        )
        .setValue("Not Sent");

      PRIVATE.applyRowColor(
        sheet,
        rowNumber,
        headers
      );

      Logger.log(
        "Initial email failed for " +
        email +
        ": " +
        error.message
      );
    }

    Utilities.sleep(1000);
  }


  /**********************************************************
   * FINAL STATUS SYNC
   **********************************************************/

  PRIVATE.syncEmailStatuses(sheet);

  Logger.log(
    "========================================"
  );

  Logger.log(
    "Initial email automation completed."
  );

  Logger.log(
    "Initial emails sent: " +
    emailsSent
  );

  Logger.log(
    "Gmail recipient quota remaining: " +
    MailApp.getRemainingDailyQuota()
  );

  Logger.log(
    "========================================"
  );
}


/************************************************************
 * PUBLIC FUNCTION 2
 * DAY 02 FOLLOW-UP
 ************************************************************/

function followUpDay02() {

  PRIVATE.processFollowUp(2);
}


/************************************************************
 * PUBLIC FUNCTION 3
 * DAY 03 FOLLOW-UP
 ************************************************************/

function followUpDay03() {

  PRIVATE.processFollowUp(3);
}


/************************************************************
 * PUBLIC FUNCTION 4
 * DAY 04 FOLLOW-UP
 ************************************************************/

function followUpDay04() {

  PRIVATE.processFollowUp(4);
}


/************************************************************
 * PRIVATE APPLICATION OBJECT
 ************************************************************/

const PRIVATE = {


  /**********************************************************
   * PROCESS FOLLOW-UP
   **********************************************************/

  processFollowUp:
    function(followUpNumber) {

      if (
        !this.isBusinessDay()
      ) {

        Logger.log(
          "Today is Saturday/Sunday. No follow-up emails will be sent."
        );

        return;
      }


      const sheet =
        SpreadsheetApp
          .getActiveSpreadsheet()
          .getActiveSheet();


      this.ensureColumns(sheet);


      // Detect replies and delivery failures before sending.
      const syncSummary =
        this.syncEmailStatuses(sheet);


      const data =
        sheet
          .getDataRange()
          .getValues();


      if (
        data.length < 2
      ) {

        Logger.log(
          "No recruiter rows found."
        );

        return;
      }


      const headers =
        data[0];


      const companyCol =
        this.findColumn(
          headers,
          ["Company"]
        );

      const emailCol =
        this.findColumn(
          headers,
          ["Email"]
        );

      const recruiterCol =
        this.findColumn(
          headers,
          ["Recruiter Name", "Recruiter"]
        );

      const statusCol =
        this.findColumn(
          headers,
          ["Status"]
        );

      const threadCol =
        this.findColumn(
          headers,
          ["Thread ID"]
        );

      const deliveryCol =
        this.findColumn(
          headers,
          ["Delivery Status"]
        );

      const responseCol =
        this.findColumn(
          headers,
          ["Response Status"]
        );

      const followStatusCol =
        this.findColumn(
          headers,
          [
            "Follow-up 0" +
            followUpNumber +
            " Status"
          ]
        );

      const followDateCol =
        this.findColumn(
          headers,
          [
            "Follow-up 0" +
            followUpNumber +
            " Date"
          ]
        );


      if (
        companyCol === -1 ||
        emailCol === -1 ||
        recruiterCol === -1 ||
        statusCol === -1 ||
        threadCol === -1 ||
        deliveryCol === -1 ||
        responseCol === -1 ||
        followStatusCol === -1 ||
        followDateCol === -1
      ) {

        throw new Error(
          "Required follow-up columns are missing."
        );
      }


      let remainingQuota =
        MailApp.getRemainingDailyQuota();


      Logger.log(
        "========================================"
      );

      Logger.log(
        "FOLLOW-UP DAY " +
        String(followUpNumber).padStart(2, "0")
      );

      Logger.log(
        "Gmail recipient quota remaining: " +
        remainingQuota
      );

      Logger.log(
        "Recruiter replies detected: " +
        syncSummary.replies
      );

      Logger.log(
        "Delivery failures detected: " +
        syncSummary.failures
      );

      Logger.log(
        "========================================"
      );


      if (
        remainingQuota <= 0
      ) {

        Logger.log(
          "No Gmail recipient quota remaining."
        );

        return;
      }


      let followUpsSent = 0;

      const today =
        new Date();


      /********************************************************
       * PROCESS EACH ROW
       ********************************************************/

      for (
        let i = 1;
        i < data.length;
        i++
      ) {

        if (
          followUpsSent >=
          CONFIG.MAX_FOLLOWUPS_PER_RUN
        ) {
          break;
        }

        if (
          remainingQuota <= 0
        ) {
          break;
        }


        const rowNumber =
          i + 1;


        const company =
          String(
            data[i][companyCol] || ""
          ).trim();


        const email =
          this.normalizeEmail(
            data[i][emailCol]
          );


        const recruiterName =
          String(
            data[i][recruiterCol] || ""
          ).trim();


        const status =
          String(
            data[i][statusCol] || ""
          )
            .trim()
            .toLowerCase();


        const threadId =
          String(
            data[i][threadCol] || ""
          ).trim();


        const deliveryStatus =
          String(
            data[i][deliveryCol] || ""
          )
            .trim()
            .toLowerCase();


        const responseStatus =
          String(
            data[i][responseCol] || ""
          )
            .trim()
            .toLowerCase();


        const followStatus =
          String(
            data[i][followStatusCol] || ""
          )
            .trim()
            .toLowerCase();


        const followDate =
          data[i][followDateCol];


        if (!email) {
          continue;
        }


        if (
          company.toUpperCase() ===
          CONFIG.TEST_COMPANY_NAME
        ) {
          continue;
        }


        if (
          status !== "sent"
        ) {
          continue;
        }


        // Stop all future follow-ups after a recruiter reply.
        if (
          responseStatus === "replied"
        ) {

          Logger.log(
            "Skipping follow-up because recruiter replied: " +
            email
          );

          continue;
        }


        // Stop all future follow-ups after bounce/block.
        if (
          deliveryStatus === "bounced" ||
          deliveryStatus === "blocked"
        ) {

          Logger.log(
            "Skipping follow-up because email failed: " +
            email
          );

          continue;
        }


        if (
          followStatus !== "pending"
        ) {
          continue;
        }


        if (!followDate) {
          continue;
        }


        if (
          this.dateOnly(
            new Date(followDate)
          ) >
          this.dateOnly(today)
        ) {

          continue;
        }


        if (
          !threadId ||
          threadId === "PENDING_RECOVERY"
        ) {

          Logger.log(
            "Skipping " +
            email +
            " because Thread ID is missing."
          );

          continue;
        }


        let thread = null;


        try {

          thread =
            GmailApp.getThreadById(
              threadId
            );

        } catch (error) {

          Logger.log(
            "Unable to open Gmail thread for " +
            email +
            ": " +
            error.message
          );

          continue;
        }


        if (!thread) {

          Logger.log(
            "Gmail thread not found for " +
            email
          );

          continue;
        }


        // Last-second reply check before sending.
        if (
          this.threadHasExternalReply(
            thread,
            email
          )
        ) {

          this.markReplied(
            sheet,
            rowNumber,
            headers,
            thread
          );

          continue;
        }


        /********************************************************
         * GET LATEST MESSAGE AS REPLY ANCHOR
         ********************************************************/

        const latestMessage =
          this.getLatestMessage(
            thread
          );


        if (!latestMessage) {

          Logger.log(
            "No latest message found for thread: " +
            threadId
          );

          continue;
        }


        const latestMessageId =
          this.getMessageHeader(
            latestMessage,
            "Message-ID"
          );


        if (!latestMessageId) {

          Logger.log(
            "Message-ID not found for thread: " +
            threadId
          );

          continue;
        }


        const existingReferences =
          this.getMessageHeader(
            latestMessage,
            "References"
          );


        const body =
          this.getFollowUpBody(
            followUpNumber,
            recruiterName || "Team"
          );


        /********************************************************
         * SEND FOLLOW-UP USING GMAIL API
         *
         * Explicit To:
         * recruiter email
         *
         * Explicit threadId:
         * existing Gmail conversation
         *
         * In-Reply-To:
         * latest message in conversation
         ********************************************************/

        try {

          const raw =
            this.buildReplyMime(
              email,
              CONFIG.EMAIL_SUBJECT,
              body,
              latestMessageId,
              existingReferences
            );


          const encoded =
            Utilities.base64EncodeWebSafe(
              raw,
              Utilities.Charset.UTF_8
            );


          Gmail.Users.Messages.send(
            {
              raw: encoded,
              threadId: threadId
            },
            "me"
          );


          sheet
            .getRange(
              rowNumber,
              followStatusCol + 1
            )
            .setValue("Sent");


          sheet
            .getRange(
              rowNumber,
              followDateCol + 1
            )
            .setValue(
              new Date()
            );


          followUpsSent++;
          remainingQuota--;


          Logger.log(
            "Day-" +
            String(
              followUpNumber
            ).padStart(2, "0") +
            " follow-up sent successfully: " +
            email +
            " | Thread ID: " +
            threadId
          );


        } catch (error) {

          Logger.log(
            "Day-" +
            String(
              followUpNumber
            ).padStart(2, "0") +
            " failed for " +
            email +
            ": " +
            error.message
          );
        }


        Utilities.sleep(1000);
      }


      this.syncEmailStatuses(sheet);


      Logger.log(
        "========================================"
      );

      Logger.log(
        "Day-" +
        String(
          followUpNumber
        ).padStart(2, "0") +
        " automation completed."
      );

      Logger.log(
        "Follow-ups sent: " +
        followUpsSent
      );

      Logger.log(
        "Gmail recipient quota remaining: " +
        MailApp.getRemainingDailyQuota()
      );

      Logger.log(
        "========================================"
      );
    },


  /**********************************************************
   * SYNCHRONIZE DELIVERY FAILURES + REPLIES
   **********************************************************/

  syncEmailStatuses:
    function(sheet) {

      this.ensureColumns(sheet);


      const data =
        sheet
          .getDataRange()
          .getValues();


      if (
        data.length < 2
      ) {

        return {
          failures: 0,
          replies: 0
        };
      }


      const headers =
        data[0];


      const emailCol =
        this.findColumn(
          headers,
          ["Email"]
        );

      const statusCol =
        this.findColumn(
          headers,
          ["Status"]
        );

      const threadCol =
        this.findColumn(
          headers,
          ["Thread ID"]
        );

      const deliveryCol =
        this.findColumn(
          headers,
          ["Delivery Status"]
        );

      const responseCol =
        this.findColumn(
          headers,
          ["Response Status"]
        );

      const responseDateCol =
        this.findColumn(
          headers,
          ["Last Response Date"]
        );

      const notesCol =
        this.findColumn(
          headers,
          ["Notes"]
        );


      if (
        emailCol === -1 ||
        statusCol === -1 ||
        threadCol === -1 ||
        deliveryCol === -1 ||
        responseCol === -1 ||
        responseDateCol === -1 ||
        notesCol === -1
      ) {

        return {
          failures: 0,
          replies: 0
        };
      }


      const failures =
        this.scanDeliveryFailures(
          sheet,
          data,
          headers,
          emailCol,
          deliveryCol,
          responseCol,
          notesCol
        );


      const refreshedData =
        sheet
          .getDataRange()
          .getValues();


      let replies = 0;


      /********************************************************
       * CHECK EACH STORED GMAIL THREAD
       ********************************************************/

      for (
        let i = 1;
        i < refreshedData.length;
        i++
      ) {

        const rowNumber =
          i + 1;


        const email =
          this.normalizeEmail(
            refreshedData[i][emailCol]
          );


        const initialStatus =
          String(
            refreshedData[i][statusCol] || ""
          )
            .trim()
            .toLowerCase();


        const threadId =
          String(
            refreshedData[i][threadCol] || ""
          ).trim();


        const deliveryStatus =
          String(
            refreshedData[i][deliveryCol] || ""
          )
            .trim()
            .toLowerCase();


        const currentResponse =
          String(
            refreshedData[i][responseCol] || ""
          )
            .trim()
            .toLowerCase();


        if (!email) {
          continue;
        }


        if (
          initialStatus !== "sent"
        ) {
          continue;
        }


        if (
          deliveryStatus === "bounced" ||
          deliveryStatus === "blocked"
        ) {
          continue;
        }


        if (
          !threadId ||
          threadId === "PENDING_RECOVERY"
        ) {
          continue;
        }


        let thread = null;


        try {

          thread =
            GmailApp.getThreadById(
              threadId
            );

        } catch (error) {

          Logger.log(
            "Unable to inspect thread " +
            threadId +
            " for " +
            email +
            ": " +
            error.message
          );

          continue;
        }


        if (!thread) {
          continue;
        }


        if (
          this.threadHasExternalReply(
            thread,
            email
          )
        ) {

          if (
            currentResponse !== "replied"
          ) {

            this.markReplied(
              sheet,
              rowNumber,
              headers,
              thread
            );

            replies++;
          }
        }
      }


      return {
        failures: failures,
        replies: replies
      };
    },


  /**********************************************************
   * SCAN GMAIL DELIVERY FAILURE NOTIFICATIONS
   **********************************************************/

  scanDeliveryFailures:
    function(
      sheet,
      data,
      headers,
      emailCol,
      deliveryCol,
      responseCol,
      notesCol
    ) {

      let failureCount = 0;


      const queries = [

        "from:mailer-daemon@googlemail.com newer_than:" +
        CONFIG.BOUNCE_LOOKBACK_DAYS +
        "d",

        "from:mailer-daemon@google.com newer_than:" +
        CONFIG.BOUNCE_LOOKBACK_DAYS +
        "d"
      ];


      const failureThreads = [];


      for (
        let q = 0;
        q < queries.length;
        q++
      ) {

        try {

          const threads =
            GmailApp.search(
              queries[q],
              0,
              CONFIG.MAX_BOUNCE_THREADS_PER_QUERY
            );


          for (
            let i = 0;
            i < threads.length;
            i++
          ) {

            failureThreads.push(
              threads[i]
            );
          }

        } catch (error) {

          Logger.log(
            "Delivery failure search failed: " +
            error.message
          );
        }
      }


      if (
        failureThreads.length === 0
      ) {

        return 0;
      }


      /********************************************************
       * MATCH EACH FAILURE AGAINST SHEET EMAILS
       ********************************************************/

      for (
        let t = 0;
        t < failureThreads.length;
        t++
      ) {

        const thread =
          failureThreads[t];


        const messages =
          thread.getMessages();


        for (
          let m = 0;
          m < messages.length;
          m++
        ) {

          const message =
            messages[m];


          const from =
            String(
              message.getFrom() || ""
            ).toLowerCase();


          if (
            from.indexOf(
              "mailer-daemon"
            ) === -1
          ) {
            continue;
          }


          const body =
            (
              String(
                message.getPlainBody() || ""
              ) +
              "\n" +
              String(
                message.getBody() || ""
              )
            ).toLowerCase();


          const subject =
            String(
              message.getSubject() || ""
            ).toLowerCase();


          if (
            body.indexOf(
              "delivery status notification"
            ) === -1 &&
            subject.indexOf(
              "delivery status notification"
            ) === -1 &&
            body.indexOf(
              "message blocked"
            ) === -1 &&
            body.indexOf(
              "address not found"
            ) === -1 &&
            body.indexOf(
              "wasn't delivered"
            ) === -1 &&
            body.indexOf(
              "was not delivered"
            ) === -1
          ) {

            continue;
          }


          let matched = false;


          for (
            let r = 1;
            r < data.length;
            r++
          ) {

            const recruiterEmail =
              this.normalizeEmail(
                data[r][emailCol]
              );


            if (!recruiterEmail) {
              continue;
            }


            if (
              body.indexOf(
                recruiterEmail
              ) === -1
            ) {
              continue;
            }


            const rowNumber =
              r + 1;


            const failureType =
              this.classifyDeliveryFailure(
                body,
                subject
              );


            const deliveryText =
              failureType === "BLOCKED"
                ? "Blocked"
                : "Bounced";


            sheet
              .getRange(
                rowNumber,
                deliveryCol + 1
              )
              .setValue(
                deliveryText
              );


            sheet
              .getRange(
                rowNumber,
                responseCol + 1
              )
              .setValue(
                "Not Delivered"
              );


            const existingNotes =
              String(
                sheet
                  .getRange(
                    rowNumber,
                    notesCol + 1
                  )
                  .getValue() || ""
              ).trim();


            const note =
              failureType === "BLOCKED"
                ? "Gmail delivery failure: Message blocked"
                : "Gmail delivery failure: Address/domain/message not delivered";


            if (
              existingNotes
                .toLowerCase()
                .indexOf(
                  note.toLowerCase()
                ) === -1
            ) {

              sheet
                .getRange(
                  rowNumber,
                  notesCol + 1
                )
                .setValue(
                  existingNotes
                    ? existingNotes +
                      " | " +
                      note
                    : note
                );
            }


            this.applyRowColor(
              sheet,
              rowNumber,
              headers
            );


            failureCount++;


            Logger.log(
              "Delivery failure detected: " +
              recruiterEmail +
              " | " +
              deliveryText
            );


            /************************************************
             * MOVE ONLY THE FAILURE NOTIFICATION TO TRASH.
             * Do NOT trash the original recruiter thread.
             ************************************************/

            try {

              message.moveToTrash();

              Logger.log(
                "Moved delivery failure notification to Trash for: " +
                recruiterEmail
              );

            } catch (trashError) {

              Logger.log(
                "Could not move failure notification to Trash: " +
                trashError.message
              );
            }


            matched = true;

            break;
          }


          if (matched) {
            continue;
          }
        }
      }


      return failureCount;
    },


  /**********************************************************
   * CLASSIFY DELIVERY FAILURE
   **********************************************************/

  classifyDeliveryFailure:
    function(
      body,
      subject
    ) {

      const combined =
        (
          body +
          " " +
          subject
        ).toLowerCase();


      if (
        combined.indexOf(
          "message blocked"
        ) !== -1 ||
        combined.indexOf(
          "has been blocked"
        ) !== -1 ||
        combined.indexOf(
          "blocked by"
        ) !== -1
      ) {

        return "BLOCKED";
      }


      return "BOUNCED";
    },


  /**********************************************************
   * DETECT EXTERNAL RECRUITER REPLY
   **********************************************************/

  threadHasExternalReply:
    function(
      thread,
      recruiterEmail
    ) {

      const ownEmail =
        this.getSenderEmail();


      const messages =
        thread.getMessages();


      for (
        let i = 0;
        i < messages.length;
        i++
      ) {

        const message =
          messages[i];


        const from =
          this.extractEmail(
            message.getFrom()
          );


        if (!from) {
          continue;
        }


        // The stored recruiter address is the strongest
        // signal for a recruiter response.
        if (
          from === recruiterEmail
        ) {

          return true;
        }


        // If Apps Script can identify our own address,
        // an incoming message from another address in the
        // same thread is treated as an external response.
        if (
          ownEmail &&
          from !== ownEmail
        ) {

          let toText = "";

          try {

            toText =
              String(
                message.getTo() || ""
              ).toLowerCase();

          } catch (error) {

            toText = "";
          }


          if (
            toText.indexOf(
              ownEmail
            ) !== -1
          ) {

            return true;
          }
        }
      }


      return false;
    },


  /**********************************************************
   * MARK REPLIED
   **********************************************************/

  markReplied:
    function(
      sheet,
      rowNumber,
      headers,
      thread
    ) {

      const responseCol =
        this.findColumn(
          headers,
          ["Response Status"]
        );


      const responseDateCol =
        this.findColumn(
          headers,
          ["Last Response Date"]
        );


      const deliveryCol =
        this.findColumn(
          headers,
          ["Delivery Status"]
        );


      const notesCol =
        this.findColumn(
          headers,
          ["Notes"]
        );


      if (
        responseCol === -1 ||
        responseDateCol === -1
      ) {
        return;
      }


      const latestExternalMessage =
        this.getLatestExternalMessage(
          thread
        );


      if (!latestExternalMessage) {
        return;
      }


      sheet
        .getRange(
          rowNumber,
          responseCol + 1
        )
        .setValue(
          "Replied"
        );


      sheet
        .getRange(
          rowNumber,
          responseDateCol + 1
        )
        .setValue(
          latestExternalMessage.getDate()
        );


      if (
        deliveryCol !== -1
      ) {

        const currentDelivery =
          String(
            sheet
              .getRange(
                rowNumber,
                deliveryCol + 1
              )
              .getValue() || ""
          )
            .trim()
            .toLowerCase();


        if (
          currentDelivery !== "bounced" &&
          currentDelivery !== "blocked"
        ) {

          sheet
            .getRange(
              rowNumber,
              deliveryCol + 1
            )
            .setValue(
              "Delivered"
            );
        }
      }


      if (
        notesCol !== -1
      ) {

        const existingNotes =
          String(
            sheet
              .getRange(
                rowNumber,
                notesCol + 1
              )
              .getValue() || ""
          ).trim();


        if (
          existingNotes
            .toLowerCase()
            .indexOf(
              "recruiter response received"
            ) === -1
        ) {

          sheet
            .getRange(
              rowNumber,
              notesCol + 1
            )
            .setValue(
              existingNotes
                ? existingNotes +
                  " | Recruiter response received"
                : "Recruiter response received"
            );
        }
      }


      this.applyRowColor(
        sheet,
        rowNumber,
        headers
      );


      Logger.log(
        "Recruiter response detected. Thread ID: " +
        thread.getId()
      );
    },


  /**********************************************************
   * GET LATEST MESSAGE
   **********************************************************/

  getLatestMessage:
    function(thread) {

      const messages =
        thread.getMessages();


      if (
        !messages ||
        messages.length === 0
      ) {

        return null;
      }


      return messages[
        messages.length - 1
      ];
    },


  /**********************************************************
   * GET LATEST EXTERNAL MESSAGE
   **********************************************************/

  getLatestExternalMessage:
    function(thread) {

      const ownEmail =
        this.getSenderEmail();


      const messages =
        thread.getMessages();


      for (
        let i = messages.length - 1;
        i >= 0;
        i--
      ) {

        const message =
          messages[i];


        const from =
          this.extractEmail(
            message.getFrom()
          );


        if (
          from &&
          (
            !ownEmail ||
            from !== ownEmail
          )
        ) {

          return message;
        }
      }


      return null;
    },


  /**********************************************************
   * GET GMAIL MESSAGE HEADER
   **********************************************************/

  getMessageHeader:
    function(
      message,
      headerName
    ) {

      try {

        const value =
          message.getHeader(
            headerName
          );

        if (value) {
          return value;
        }

      } catch (error) {
        // Fall back to Gmail API.
      }


      try {

        const apiMessage =
          Gmail.Users.Messages.get(
            "me",
            message.getId(),
            {
              format: "metadata",
              metadataHeaders: [
                headerName
              ]
            }
          );


        const headers =
          (
            apiMessage.payload &&
            apiMessage.payload.headers
          ) || [];


        for (
          let i = 0;
          i < headers.length;
          i++
        ) {

          if (
            String(
              headers[i].name
            ).toLowerCase() ===
            headerName.toLowerCase()
          ) {

            return headers[i].value;
          }
        }

      } catch (error) {

        Logger.log(
          "Unable to retrieve Gmail header " +
          headerName +
          ": " +
          error.message
        );
      }


      return "";
    },


  /**********************************************************
   * BUILD REPLY MIME MESSAGE
   **********************************************************/

  buildReplyMime:
    function(
      to,
      subject,
      body,
      inReplyTo,
      existingReferences
    ) {

      const references =
        existingReferences
          ? existingReferences +
            " " +
            inReplyTo
          : inReplyTo;


      return (
        "To: " +
        to +
        "\r\n" +

        "Subject: " +
        subject +
        "\r\n" +

        "MIME-Version: 1.0\r\n" +

        "Content-Type: text/plain; charset=UTF-8\r\n" +

        "Content-Transfer-Encoding: 8bit\r\n" +

        "In-Reply-To: " +
        inReplyTo +
        "\r\n" +

        "References: " +
        references +
        "\r\n\r\n" +

        body
      );
    },


  /**********************************************************
   * FOLLOW-UP EMAIL CONTENT
   **********************************************************/

  getFollowUpBody:
    function(
      followUpNumber,
      name
    ) {

      if (
        followUpNumber === 2
      ) {

        return `Hi ${name},

I hope you’re doing well.

I wanted to follow up on my application and briefly highlight the technical experience I can bring to the above-mentioned role.

My recent experience has involved building and maintaining automation frameworks using Selenium WebDriver, Java, TestNG, Cucumber BDD, POM, Maven, and Git. Along with UI automation, I have worked on API validation and automation using Rest Assured and Postman, as well as CI/CD based regression execution through Jenkins.

I have also upgraded my skills with Playwright and have experience covering web, API, and end-to-end business workflows. My focus has been on developing reusable automation components and maintaining reliable regression suites rather than simply creating individual test scripts.

I would be glad to discuss my technical experience in more detail if my profile is relevant to the current requirement.

Please let me know if there is any update regarding my application or the next step.

Regards,
Harshad Chandrakant Shinde`;
      }


      if (
        followUpNumber === 3
      ) {

        return `Hi ${name},

I hope you’re doing well.

I wanted to check in regarding my application and share a little more about the kind of impact I have delivered in my recent QA role.

At Xenneo Tech, I worked on a project where I was involved in automation framework development, API/UI automation, regression testing, and release validation. The automation framework I worked on helped improve regression efficiency by around 60%, while the automation coverage was expanded to more than 450 test scenarios. I also contributed to reducing manual testing effort by approximately 70%.

My experience has involved testing business-critical workflows such as booking, onboarding, payment, cart, and checkout, along with API and backend validation.

I’m particularly interested in opportunities where I can contribute not only as an automation tester but also in improving test coverage, regression efficiency, and overall release quality.

Please let me know if there is any update on my candidature. I would be happy to discuss my experience further.

Regards,
Harshad Chandrakant Shinde`;
      }


      if (
        followUpNumber === 4
      ) {

        return `Hi ${name},

I hope you’re doing well.

I’m following up once again regarding my application for the above-mentioned opportunity. I wanted to check whether my profile is currently being considered and if there are any updates regarding the next stage of the recruitment process.

I have 5 years 7 months of QA experience with hands-on expertise in automation, API testing, framework development, and Agile delivery. I have also completed my notice period and am available to join immediately.

If the position is still open, I'm happy to participate in the interview process at your convenience. If my profile isn't a fit for this requirement, please let me know if any other QA Automation openings would be a better fit.

Thank you for your time and consideration.

Regards,
Harshad Chandrakant Shinde`;
      }


      throw new Error(
        "Invalid follow-up number: " +
        followUpNumber
      );
    },


  /**********************************************************
   * ENSURE REQUIRED COLUMNS
   **********************************************************/

  ensureColumns:
    function(sheet) {

      const requiredColumns = [

        "Company",
        "Email",
        "Recruiter Name",

        "Status",
        "Delivery Status",
        "Response Status",

        "Sent Date",
        "Last Response Date",

        "Follow-up 02 Status",
        "Follow-up 02 Date",

        "Follow-up 03 Status",
        "Follow-up 03 Date",

        "Follow-up 04 Status",
        "Follow-up 04 Date",

        "Thread ID",
        "Notes"
      ];


      let lastColumn =
        sheet.getLastColumn();


      if (
        lastColumn === 0
      ) {

        sheet
          .getRange(1, 1)
          .setValue("Company");

        lastColumn = 1;
      }


      let headers =
        sheet
          .getRange(
            1,
            1,
            1,
            lastColumn
          )
          .getValues()[0];


      requiredColumns.forEach(
        function(columnName) {

          let exists = false;


          for (
            let i = 0;
            i < headers.length;
            i++
          ) {

            if (
              String(
                headers[i] || ""
              )
                .trim()
                .toLowerCase() ===
              columnName.toLowerCase()
            ) {

              exists = true;
              break;
            }
          }


          if (!exists) {

            lastColumn++;


            sheet
              .getRange(
                1,
                lastColumn
              )
              .setValue(
                columnName
              );


            headers.push(
              columnName
            );
          }
        }
      );


      sheet
        .getRange(
          1,
          1,
          1,
          lastColumn
        )
        .setFontWeight("bold");
    },


  /**********************************************************
   * APPLY COLOR
   *
   * RED   = bounced / blocked
   * BLUE  = recruiter replied
   * GREEN = sent and no bounce detected
   * YELLOW= pending / waiting
   **********************************************************/

  applyRowColor:
    function(
      sheet,
      rowNumber,
      headers
    ) {

      const emailCol =
        this.findColumn(
          headers,
          ["Email"]
        );

      const statusCol =
        this.findColumn(
          headers,
          ["Status"]
        );

      const deliveryCol =
        this.findColumn(
          headers,
          ["Delivery Status"]
        );

      const responseCol =
        this.findColumn(
          headers,
          ["Response Status"]
        );


      if (
        emailCol === -1
      ) {
        return;
      }


      const status =
        String(
          sheet
            .getRange(
              rowNumber,
              statusCol + 1
            )
            .getValue() || ""
        )
          .trim()
          .toLowerCase();


      const delivery =
        String(
          sheet
            .getRange(
              rowNumber,
              deliveryCol + 1
            )
            .getValue() || ""
        )
          .trim()
          .toLowerCase();


      const response =
        String(
          sheet
            .getRange(
              rowNumber,
              responseCol + 1
            )
            .getValue() || ""
        )
          .trim()
          .toLowerCase();


      let color =
        CONFIG.COLORS.YELLOW;


      if (
        delivery === "bounced" ||
        delivery === "blocked"
      ) {

        color =
          CONFIG.COLORS.RED;

      } else if (
        response === "replied"
      ) {

        color =
          CONFIG.COLORS.BLUE;

      } else if (
        delivery === "delivered" &&
        status === "sent"
      ) {

        color =
          CONFIG.COLORS.GREEN;

      } else if (
        status === "sent"
      ) {

        color =
          CONFIG.COLORS.GREEN;
      }


      // Primary visual indicator: Email cell.
      sheet
        .getRange(
          rowNumber,
          emailCol + 1
        )
        .setBackground(color);


      // Delivery Status cell.
      if (
        deliveryCol !== -1
      ) {

        sheet
          .getRange(
            rowNumber,
            deliveryCol + 1
          )
          .setBackground(
            delivery === "bounced" ||
            delivery === "blocked"
              ? CONFIG.COLORS.RED
              : delivery === "delivered"
                ? CONFIG.COLORS.GREEN
                : CONFIG.COLORS.YELLOW
          );
      }


      // Response Status cell.
      if (
        responseCol !== -1
      ) {

        sheet
          .getRange(
            rowNumber,
            responseCol + 1
          )
          .setBackground(
            response === "replied"
              ? CONFIG.COLORS.BLUE
              : response === "waiting"
                ? CONFIG.COLORS.YELLOW
                : response === "not delivered"
                  ? CONFIG.COLORS.RED
                  : CONFIG.COLORS.YELLOW
          );
      }
    },


  /**********************************************************
   * FIND COLUMN
   **********************************************************/

  findColumn:
    function(
      headers,
      possibleNames
    ) {

      for (
        let i = 0;
        i < headers.length;
        i++
      ) {

        const header =
          String(
            headers[i] || ""
          )
            .trim()
            .toLowerCase();


        for (
          let j = 0;
          j < possibleNames.length;
          j++
        ) {

          if (
            header ===
            possibleNames[j]
              .toLowerCase()
          ) {

            return i;
          }
        }
      }


      return -1;
    },


  /**********************************************************
   * NORMALIZE EMAIL
   **********************************************************/

  normalizeEmail:
    function(value) {

      return String(
        value || ""
      )
        .trim()
        .toLowerCase();
    },


  /**********************************************************
   * EXTRACT EMAIL FROM:
   * John Doe <john@example.com>
   **********************************************************/

  extractEmail:
    function(value) {

      const text =
        String(
          value || ""
        )
          .trim()
          .toLowerCase();


      const match =
        text.match(
          /<([^>]+)>/
        );


      if (
        match &&
        match[1]
      ) {

        return this.normalizeEmail(
          match[1]
        );
      }


      const emailMatch =
        text.match(
          /[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i
        );


      return emailMatch
        ? this.normalizeEmail(
            emailMatch[0]
          )
        : "";
    },


  /**********************************************************
   * GET SENDER EMAIL
   **********************************************************/

  getSenderEmail:
    function() {

      if (
        CONFIG.SENDER_EMAIL
      ) {

        return this.normalizeEmail(
          CONFIG.SENDER_EMAIL
        );
      }


      try {

        const effective =
          Session
            .getEffectiveUser()
            .getEmail();


        if (effective) {

          return this.normalizeEmail(
            effective
          );
        }

      } catch (error) {
        // Continue.
      }


      try {

        const active =
          Session
            .getActiveUser()
            .getEmail();


        if (active) {

          return this.normalizeEmail(
            active
          );
        }

      } catch (error) {
        // Continue.
      }


      return "";
    },


  /**********************************************************
   * DERIVE RECRUITER NAME
   **********************************************************/

  deriveRecruiterName:
    function(email) {

      const local =
        this.normalizeEmail(
          email
        )
          .split("@")[0];


      if (!local) {
        return "Team";
      }


      const genericMailboxes = [

        "hr",
        "hiring",
        "recruitment",
        "recruiter",
        "careers",
        "career",
        "jobs",
        "job",
        "talent",
        "talentacquisition",
        "sourcing",
        "source",
        "info",
        "admin",
        "support",
        "contact",
        "team",
        "hello",
        "office",
        "ops",
        "operations",
        "enquiry",
        "inquiry"
      ];


      if (
        genericMailboxes.indexOf(
          local
        ) !== -1
      ) {

        return "Team";
      }


      const cleaned =
        local
          .replace(
            /[0-9]+$/g,
            ""
          )
          .replace(
            /[._-]+/g,
            " "
          )
          .trim();


      if (!cleaned) {
        return "Team";
      }


      return cleaned
        .split(/\s+/)
        .map(
          function(part) {

            return (
              part
                .charAt(0)
                .toUpperCase() +
              part.slice(1)
            );
          }
        )
        .join(" ");
    },


  /**********************************************************
   * MONDAY-FRIDAY CHECK
   **********************************************************/

  isBusinessDay:
    function() {

      const day =
        Number(
          Utilities.formatDate(
            new Date(),
            CONFIG.TIMEZONE,
            "u"
          )
        );


      return (
        day >= 1 &&
        day <= 5
      );
    },


  /**********************************************************
   * ADD BUSINESS DAYS
   **********************************************************/

  addBusinessDays:
    function(
      startDate,
      numberOfDays
    ) {

      const result =
        new Date(
          startDate.getFullYear(),
          startDate.getMonth(),
          startDate.getDate(),
          12,
          0,
          0
        );


      let added = 0;


      while (
        added < numberOfDays
      ) {

        result.setDate(
          result.getDate() + 1
        );


        const day =
          result.getDay();


        if (
          day >= 1 &&
          day <= 5
        ) {

          added++;
        }
      }


      return result;
    },


  /**********************************************************
   * DATE WITHOUT TIME
   **********************************************************/

  dateOnly:
    function(date) {

      return new Date(
        date.getFullYear(),
        date.getMonth(),
        date.getDate()
      ).getTime();
    }
};
