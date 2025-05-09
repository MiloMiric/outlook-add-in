Office.onReady(() => {
  // Wait 500ms to make sure the item is fully loaded (you can adjust this if needed)
  setTimeout(run, 200);
});

function run() {
  const item = Office.context.mailbox.item;

  if (!item || !item.body) {
    console.error("Mailbox item is not available.");
    return;
  }

  // Function to format date as MM/DD/YYYY hh:mm
  function formatDate(date) {
    const pad = (n) => (n < 10 ? '0' + n : n);
    const month = pad(date.getMonth() + 1);
    const day = pad(date.getDate());
    const year = date.getFullYear();
    const hours = pad(date.getHours());
    const minutes = pad(date.getMinutes());
    return `${month}/${day}/${year} ${hours}:${minutes} Central Time`;
  }

  const now = new Date();
  const checkIn = formatDate(now);
  const checkOut = formatDate(new Date(now.getTime() + 3 * 60 * 60 * 1000)); // +3 hours
  const rcvrCheckIn = formatDate(new Date(now.getTime() + 6 * 60 * 60 * 1000)); // +6 hours

  const htmlTable = `
    <p style="font-family: Arial, sans-serif;">Load details:</p>
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
      <tr><td style="padding: 6px 10px; border: 1px solid #000;">Shipper Check-In</td><td style="padding: 6px 10px; border: 1px solid #000;">${checkIn}</td></tr>
      <tr><td style="padding: 6px 10px; border: 1px solid #000;">Shipper Check-Out</td><td style="padding: 6px 10px; border: 1px solid #000;">${checkOut}</td></tr>
      <tr><td style="padding: 6px 10px; border: 1px solid #000;">RCVR Check-In</td><td style="padding: 6px 10px; border: 1px solid #000;">${rcvrCheckIn}</td></tr>
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
