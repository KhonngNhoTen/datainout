import * as path from "path";
import { Reporter } from "./reports/Reporter";
import { ReportData } from "./reports/type";

async function main() {
  const reporter = new Reporter({ exporterType: "excel", templatePath: path.join(__dirname, "../templates/test.template.js") });
  const data: ReportData = {
    header: {
      title: "TITLE NEW ADADADS",
      totalOrder: 1,
      totalSuccessfulOrder: 1,
      successRate: 1,
      totalTransactionValue: 1,
      totalCancelledOrder: 1,
    },
    footer: {
      done: "DONE----",
      done2: "DONE222",
    },
    table: [
      {
        index: 1,
        id: "aa",
        bookingCode: "SSS",
        createdAt: "2023-12-12",
        deliveryAt: "2023-12-12",
        deliveryStatus: "Cancelled",
        passengerName: "AAAAAAA",
        phone: "02313213212",
        orderValue: 1,
        shippingFee: 1,
        totalPayment: 1,
        VAT: 1,
        note: "NOTE?>>",
      },
      {
        index: 2,
        id: "aa",
        bookingCode: "SSS",
        createdAt: "2023-12-12",
        deliveryAt: "2023-12-12",
        deliveryStatus: "Cancelled",
        passengerName: "AAAAAAA",
        phone: "02313213212",
        orderValue: 1,
        shippingFee: 1,
        totalPayment: 1,
        VAT: 1,
        note: "NOTE?>>",
      },
      {
        index: 3,
        id: "aAASa",
        bookingCode: "SSS",
        createdAt: "2023-12-12",
        deliveryAt: "2023-12-12",
        deliveryStatus: "Cancelled",
        passengerName: "AAAAAAA",
        phone: "02313213212",
        orderValue: 1,
        shippingFee: 1,
        totalPayment: 1,
        VAT: 1,
        note: "NOTE?>>",
      },
    ],
  };
  await reporter.writeFile(data, path.join(__dirname, "../reports-output/output.xlsx"));
}

main();
