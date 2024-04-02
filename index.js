import express from "express";
import bodyParser from "body-parser";
import http from "http";
import { invoiceReportMTD_CH } from "./src/services/DatabaseService.js";

const app = express();
const server = http.createServer(app);

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

app.get("/", function (req, res) {
  res.sendFile("/index.html", { root: "." });
});

app.get("/create", async function (req, res) {
  try {
    await invoiceReportMTD_CH();
    res.send("Invoice report generated successfully");
  } catch (error) {
    console.error(error);
    res.status(500).send("Error generating report");
  }
});

app.set("port", process.env.PORT || 3000);
server.listen(app.get("port"), function () {
  console.log("listening on port", app.get("port"));
});
