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
<p> Dear customer, <p/>
    <div style="font-family: Arial, sans-serif; max-width: 420px; margin: 15px auto; border: 1px solid #ddd; border-radius: 8px; box-shadow: 0 3px 6px rgba(0,0,0,0.1); overflow: hidden;">
      <div style="background-color: #007BFF; color: white; padding: 10px; text-align: center; font-size: 14px; font-weight: bold;">
        📦 Load Details Summary
      </div>
      <table style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 10px; text-align: center;">
        <thead>
          <tr style="background-color: #f1f1f1;">
            <th style="padding: 8px; border: 1px solid #ddd;">Item</th>
            <th style="padding: 8px; border: 1px solid #ddd;">Details</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="padding: 7px; border: 1px solid #ddd;">Load #</td>
            <td style="padding: 7px; border: 1px solid #ddd;">0158770</td>
          </tr>
          <tr>
            <td style="padding: 7px; border: 1px solid #ddd;">BOL#</td>
            <td style="padding: 7px; border: 1px solid #ddd;">2862744</td>
          </tr>
          <tr>
            <td style="padding: 7px; border: 1px solid #ddd;">Cust Ref#</td>
            <td style="padding: 7px; border: 1px solid #ddd;">2862744</td>
          </tr>
          <tr>
            <td style="padding: 7px; border: 1px solid #ddd;">Route</td>
            <td style="padding: 7px; border: 1px solid #ddd;">Mogadore, OH → Amanda, OH</td>
          </tr>
          <tr>
            <td style="padding: 7px; border: 1px solid #ddd;">Status</td>
            <td style="padding: 7px; border: 1px solid #ddd;">Arrived at shipper, still not loaded, in dock 8</td>
          </tr>
          <tr>
            <td style="padding: 7px; border: 1px solid #ddd;">Shipper Check-In</td>
            <td style="padding: 7px; border: 1px solid #ddd;">${checkIn}</td>
          </tr>
          <tr>
            <td style="padding: 7px; border: 1px solid #ddd;">Shipper Check-Out</td>
            <td style="padding: 7px; border: 1px solid #ddd;">${checkOut}</td>
          </tr>
          <tr>
            <td style="padding: 7px; border: 1px solid #ddd;">RCVR Check-In</td>
            <td style="padding: 7px; border: 1px solid #ddd;">${rcvrCheckIn}</td>
          </tr>
          <tr>
            <td style="padding: 7px; border: 1px solid #ddd;">RCVR Check-Out</td>
            <td style="padding: 7px; border: 1px solid #ddd;">Pending</td>
          </tr>
        </tbody>
      </table>
    </div>
  `;

  item.body.setSelectedDataAsync(htmlTable, { coercionType: Office.CoercionType.Html }, result => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("Insert error:", result.error.message);
    } else {
      console.log("Table inserted successfully!");
    }
  });
}
