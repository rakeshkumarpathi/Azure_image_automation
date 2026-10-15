
import os
from pathlib import Path

from dotenv import load_dotenv
from azure.storage.blob import (
    BlobServiceClient,
    ContentSettings,
)

# ==========================================
# 1. PROJECT AND SOURCE FOLDERS
# ==========================================

PROJECT_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_DIR / "data_collection"

# Check these folder names against your actual folders.
FOLDERS_TO_UPLOAD = {
    "Hardware": "Hardware",
    "Sanitory": "Sanitory",
    "Electrical":"Electrical"
    }

# ==========================================
# 2. AZURE SETTINGS
# ==========================================

CONTAINER_NAME = "nivora"

AZURE_BASE_FOLDER = "Nivora Products/categories"

IMAGE_TYPES = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".webp": "image/webp",
}

# ==========================================
# 3. LOAD AZURE CONNECTION STRING
# ==========================================

load_dotenv(PROJECT_DIR / ".env")

CONNECTION_STRING = os.getenv(
    "AZURE_STORAGE_CONNECTION_STRING"
)

if not CONNECTION_STRING:
    raise ValueError(
        "Azure connection string not found. "
        "Check the .env file in your project folder."
    )

# ==========================================
# 4. VERIFY SOURCE FOLDERS
# ==========================================

if not DATA_DIR.is_dir():
    raise FileNotFoundError(
        f"Data folder not found: {DATA_DIR}"
    )

selected_folders = {}

for category, folder_name in FOLDERS_TO_UPLOAD.items():
    source = DATA_DIR / folder_name

    if not source.is_dir():
        raise FileNotFoundError(
            f"Folder not found: {source}\n"
            "Check the actual folder name and spelling."
        )

    selected_folders[category] = source

# ==========================================
# 5. CONNECT TO AZURE
# ==========================================

service = BlobServiceClient.from_connection_string(
    CONNECTION_STRING
)

container = service.get_container_client(
    CONTAINER_NAME
)

print("Azure connection successful!")
print("Container:", CONTAINER_NAME)

# ==========================================
# 6. UPLOAD ONLY HARDWARE AND SANITORY
# ==========================================

uploaded = 0
skipped = 0

for category, source_folder in selected_folders.items():

    print(f"\nProcessing category: {category}")

    category_uploaded = 0

    # Recursively upload image files only.
    for file_path in source_folder.rglob("*"):

        if not file_path.is_file():
            continue

        extension = file_path.suffix.lower()

        if extension not in IMAGE_TYPES:
            skipped += 1
            continue

        relative_path = file_path.relative_to(source_folder)

        blob_name = (
            f"{AZURE_BASE_FOLDER}/{category}/"
            f"{relative_path.as_posix()}"
        )

        content_type = IMAGE_TYPES[extension]

        print(f"Uploading: {relative_path.as_posix()}")

        with file_path.open("rb") as file_data:
            container.upload_blob(
                name=blob_name,
                data=file_data,
                overwrite=True,
                content_settings=ContentSettings(
                    content_type=content_type
                ),
            )

        uploaded += 1
        category_uploaded += 1

    print(
        f"{category} completed: "
        f"{category_uploaded} images uploaded"
    )

# ==========================================
# 7. UPLOAD SUMMARY
# ==========================================

print("\n====================================")
print("HARDWARE AND SANITORY UPLOAD FINISHED")
print(f"Images uploaded: {uploaded}")
print(f"Non-image files skipped: {skipped}")
print("====================================")

print(
    f"\nAzure paths:\n"
    f"{CONTAINER_NAME}/{AZURE_BASE_FOLDER}/Hardware/\n"
    f"{CONTAINER_NAME}/{AZURE_BASE_FOLDER}/Sanitory/"
)
