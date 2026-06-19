const form = document.querySelector("#upload-form");
const statusEl = document.querySelector("#status");
const previewButton = document.querySelector("#preview-button");
const reviewSummary = document.querySelector("#review-summary");
const warningList = document.querySelector("#warning-list");

function filenameFromDisposition(disposition) {
  const starMatch = disposition.match(/filename\*=utf-8''([^;]+)/i);
  if (starMatch) return decodeURIComponent(starMatch[1]);
  const match = disposition.match(/filename="?([^";]+)"?/i);
  if (match) return decodeURIComponent(match[1]);
  return "Quote Comparison.xlsx";
}

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  statusEl.textContent = "Generating workbook...";

  const existing = document.querySelector("#existing-comparison").files[0];
  const data = new FormData();
  let endpoint = "/api/comparisons/export";

  if (existing) {
    endpoint = "/api/comparisons/update";
    const revised = document.querySelector("#vendors").files[0];
    data.append("existing_comparison", existing);
    data.append("revised_vendor", revised);
    data.append("vendor_name", document.querySelector("#revised-vendor-name").value);
    data.append("revision", document.querySelector("#revision").value);
  } else {
    data.append("original", document.querySelector("#original").files[0]);
    for (const file of document.querySelector("#vendors").files) {
      data.append("vendors", file);
    }
    data.append("vendor_names", document.querySelector("#vendor-names").value);
  }

  const response = await fetch(endpoint, {
    method: "POST",
    body: data,
  });

  if (!response.ok) {
    statusEl.textContent = "Export failed. Check the uploaded files and try again.";
    return;
  }

  const blob = await response.blob();
  const filename = filenameFromDisposition(response.headers.get("content-disposition") || "");
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
  statusEl.textContent = `Downloaded ${filename}`;
});

previewButton.addEventListener("click", async () => {
  statusEl.textContent = "Preparing review...";
  const existing = document.querySelector("#existing-comparison").files[0];
  const data = new FormData();
  let endpoint = "/api/comparisons/preview";

  if (existing) {
    endpoint = "/api/comparisons/update/preview";
    const revised = document.querySelector("#vendors").files[0];
    data.append("existing_comparison", existing);
    data.append("revised_vendor", revised);
    data.append("vendor_name", document.querySelector("#revised-vendor-name").value);
    data.append("revision", document.querySelector("#revision").value);
  } else {
    data.append("original", document.querySelector("#original").files[0]);
    for (const file of document.querySelector("#vendors").files) {
      data.append("vendors", file);
    }
    data.append("vendor_names", document.querySelector("#vendor-names").value);
  }

  const response = await fetch(endpoint, {
    method: "POST",
    body: data,
  });

  if (!response.ok) {
    statusEl.textContent = "Preview failed. Check the uploaded files and try again.";
    return;
  }

  const payload = await response.json();
  warningList.replaceChildren();
  reviewSummary.textContent = `${payload.project_name} - ${payload.vendors.join(", ")} - ${payload.warnings.length} warning(s)`;
  for (const warning of payload.warnings.slice(0, 50)) {
    const item = document.createElement("li");
    item.textContent = `${warning.vendor_name || "General"} | ${warning.package || ""} | ${warning.code}: ${warning.message}`;
    warningList.appendChild(item);
  }
  statusEl.textContent = "Review ready.";
});
