import sgMail from "@sendgrid/mail";
sgMail.setApiKey(process.env.SENDGRID_API_KEY);

async function sendEmailWithAttachment(
  buffer,
  fileName,
  recipients,
  subject,
  textContent,
) {

  const msg = {
    to: recipients,
    from: "notifications@amtauto.co.uk",
    subject: subject,
    text: textContent,
    attachments: [
      {
        filename: fileName,
        content: buffer.toString('base64'),
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disposition: "attachment",
      },
    ],
  };


  await sgMail.send(msg);
}

export { sendEmailWithAttachment };
