function generateEmailBody(options) {
  const {
    internName,
    internshipName,
    cohortNumber,
    deadline,
    missingItemsHtml,
    yourName 
  } = options;

  return `
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
        p { margin-bottom: 10px; }
        ul { margin-top: 5px; margin-bottom: 15px; padding-left: 20px; }
        li { margin-bottom: 5px; }
        b { font-weight: bold; }
        u { text-decoration: underline; }
        a { color: #1a73e8; text-decoration: none; } 
        a:hover { text-decoration: underline; }
        .signature { color: gray; margin-top: 20px; padding-top: 10px; border-top: 1px solid #eee; }
        .signature p { margin-bottom: 0; }
        .signature img { display: block; margin-top: 1px; }
      </style>
    </head>
    <body>
      <p>Dear <b>${internName}</b>,</p>
      <p>We hope this email finds you well.</p>
      <p>This is to inform you that for the <b>${internshipName} Virtual Early Internship</b> that you began with us on <b>${cohortNumber} 2025</b>, we have not received the following required submissions from you:</p>
      <ul>
        ${missingItemsHtml}
      </ul>
      <p><b>Please make sure to submit all the above submissions by <b>${deadline} before 11:59 PM IST</b>. Failure to submit these items will result in ineligibility for the internship.</b></p>
      <p>If you have already submitted any of these items, kindly reply to this email so we can verify and update your status accordingly.<u> Make sure that you are submitting the forms through your registered email address only.</u></p>
      <p>Should you have any questions or concerns, please do not hesitate to reach out to us.</p>
      <p> ---- </p>
      <div class="signature">
        <p style="margin-bottom: 0;">Best Regards,<br><br><b>${yourName}</b><br>Talent Discovery Team</p>
        <a href="https://4excelerate.org/" style="display: block; margin-top: 1px;"><img src="cid:signatureImage" width="150px"></a>
      </div>
    </body>
    </html>
  `;
}
