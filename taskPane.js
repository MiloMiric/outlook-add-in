Office.onReady(() => {
  insertTable(); // Odmah poziva funkciju čim se sidebar učita
});

function insertTable() {
  const item = Office.context.mailbox.item;

  const htmlTable = `
    <p style="font-family: Arial, sans-serif;">Load Details:</p>
    <table style="border-collapse: collapse; font-family: Arial, sans-serif; font-size: 13px;">
      <tr style="background-color: #cce5ff;">
        <th style="padding: 6px 10px; border: 1px solid #000; white-space: nowrap;">Item</th>
        <th style="padding: 6px 10px; border: 1px solid #000;">Details</th>
      </tr>
      <tr><td style="padding: 6px 10px; border: 1px solid #000;">Load #</td><td style="padding: 6px 10px; border: 1px solid #000;">0158770</td></tr>
      <tr><td style="padding: 6px 10px; border: 1px solid #000;">BOL#</td><td style="padding: 6px 10px; border: 1px solid #000;">2862744</td></tr>
      <tr><td style="padding: 6px 10px; border: 1px solid #000;">Cust Ref#</td><td style="padding: 6px 10px; border: 1px solid #000;">2862744</td></tr>
      <tr><td style="padding: 6px 10px; border: 1px solid #000;">Route</td><td style="padding: 6px 10px; border: 1px solid #000;">Mogadore, OH → Amanda, OH</td></tr>
      <tr><td style="padding: 6px 10px; border: 1px solid #000;">Status</td><td style="padding: 6px 10px; border: 1px solid #000;">Arrived at shipper, still not loaded, in dock 8</td></tr>
      <tr><td style="padding: 6px 10px; border: 1px solid #000;">Shipper Check-In</td><td style="padding: 6px 10px; border: 1px solid #000;">05/06/2025 09:00 Central Time</td></tr>
      <tr><td style="padding: 6px 10px; border: 1px solid #000;">Shipper Check-Out</td><td style="padding: 6px 10px; border: 1px solid #000;">05/06/2025 12:15 Central Time</td></tr>
      <tr><td style="padding: 6px 10px; border: 1px solid #000;">RCVR Check-In</td><td style="padding: 6px 10px; border: 1px solid #000;">05/06/2025 15:45 Central Time</td></tr>
      <tr><td style="padding: 6px 10px; border: 1px solid #000;">RCVR Check-Out</td><td style="padding: 6px 10px; border: 1px solid #000;">Pending</td></tr>
    </table>
  `;

  item.body.setSelectedDataAsync(htmlTable, { coercionType: Office.CoercionType.Html }, result => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("Insert error:", result.error.message);
    } else {
      console.log("Table inserted successfully!");
    }
  });
}
