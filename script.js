document.getElementById("uploadForm").addEventListener("submit", async function (e) {
    e.preventDefault();

    const fileInput = document.getElementById("fileInput");
    if (!fileInput.files.length) {
        alert("Please select at least one image to upload.");
        return;
    }

    const formData = new FormData();
    // Append all selected files
    for (let i = 0; i < fileInput.files.length; i++) {
        formData.append("files", fileInput.files[i]);
    }

    // Show progress indicator
    const progress = document.getElementById("progress");
    progress.classList.remove("hidden");

    try {
        const response = await fetch("/upload", {
            method: "POST",
            body: formData
        });

        if (response.ok) {
            // Hide progress indicator
            progress.classList.add("hidden");

            // Show download section
            const downloadSection = document.getElementById("downloadSection");
            downloadSection.classList.remove("hidden");

            // Create download link
            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const downloadLink = document.getElementById("downloadLink");
            downloadLink.href = url;
        } else {
            throw new Error("Upload failed with status: " + response.status);
        }
    } catch (error) {
        console.error("Error during upload:", error);
        alert("Failed to process your images. Please try again later.");
        progress.classList.add("hidden");
    }
});