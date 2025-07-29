async function visualizeContent() {
  try {
    const pageContent = await OneNote.run(async (context) => {
      const page = context.application.getActivePage();
      page.load("content");
      await context.sync();
      return page.content;
    });

    // Parse and visualize the content
    const container = document.getElementById("visualizationContainer");
    container.innerHTML = ""; // Clear previous visualizations

    // Example: Use D3.js to create a simple visualization
    const data = parseContentToData(pageContent);
    createVisualization(data, container);
  } catch (error) {
    console.error("Error visualizing content:", error);
  }
}

function parseContentToData(content) {
  // Parse OneNote page content into a data structure for visualization
  // Example: Extract headings, paragraphs, etc.
  return [
    { type: "heading", text: "Example Heading" },
    { type: "paragraph", text: "Example paragraph content." },
  ];
}

function createVisualization(data, container) {
  // Use a library like D3.js or Chart.js to create visualizations
  data.forEach((item) => {
    const div = document.createElement("div");
    div.style.border = "1px solid #ccc";
    div.style.margin = "10px";
    div.style.padding = "10px";
    div.textContent = `${item.type}: ${item.text}`;
    container.appendChild(div);
  });
}