
# PowerPoint Image Generation and Text-to-Speech Script

## Download the Siri Shortcut for removing background from images
Download the Siri shortcut for removing background from images. Run it manually the first time to accept permissions. 
[Download Siri Shortcut](https://www.icloud.com/shortcuts/2c1e6bc659c748a494b06f7644987972)

## Purpose
This script automates the process of adding images and text-to-speech (TTS) audio to PowerPoint presentations. It performs the following key tasks:
1. Extracts text from PowerPoint slides.
2. Uses a language model (LM Studio) to generate image descriptions based on the slide content.
3. Uses DrawThings Mac app to generate images based on the descriptions.
4. Optionally removes the background from the generated images using a Siri shortcut.
5. Inserts both the original and modified images (if available) into the PowerPoint slides.
6. Converts slide text to audio and adds the audio to the PowerPoint slides.

## Requirements
- **LM Studio**: Used to generate image descriptions. You need LM Studio running locally.
- **DrawThings**: An image generation tool based on Stable Diffusion models, which works with the script for generating images based on descriptions. You need DrawThings installed and running.
- **Siri Shortcuts**: Used to automate background removal for generated images. A predefined Siri Shortcut (RemoveBackground) is triggered from within the script.

## Key Variables to Adjust
- **LM_STUDIO_URL**: The local URL for your LM Studio instance. Default is "http://localhost:1234/v1/chat/completions".
- **DRAW_THINGS_URL**: The local URL for your DrawThings API instance. Default is "http://127.0.0.1:7860/sdapi/v1/img2img".
- **STEPS**: Number of steps for the image generation process in DrawThings. The default is set to 4.
- **REMOVE_BACKGROUND_SHORTCUT**: The name of the Siri Shortcut used for background removal. Default is "RemoveBackground".
- **addOriginal**: If set to True, the original image will be added to the PowerPoint slides. Default is False.
- **addSeparated**: If set to True, the background-removed image will be added to the PowerPoint slides. Default is True.

## Script Flow
1. **Extract Text from PowerPoint Slides**: The script extracts all text from the selected PowerPoint file using the `extract_text_from_pptx` function.
2. **Generate Image Descriptions**: It sends the extracted text of each slide to LM Studio, requesting a description of the slide’s context. The response from LM Studio will be a short description used for generating images.
3. **Generate Image**: The description is sent to DrawThings, which uses Stable Diffusion models to generate an image. You can adjust the number of steps in the generation process (default is 4). The generated image is saved to a temporary path.
4. **Background Removal**: After generating the image, the script uses a Siri Shortcut (called RemoveBackground) to remove the image’s background. The background-removed image will be downloaded into the default Downloads folder.
5. **Insert Image into PowerPoint**: The original image is inserted into the slide. If the background removal is successful and addSeparated is enabled, the background-removed image will also be added.
6. **Text-to-Speech**: The text from each slide is converted into audio using macOS’ say command, and the generated audio is inserted into the slide.
7. **Save PowerPoint**: After processing all slides, the PowerPoint presentation is saved with the newly generated images and audio.

## Setup Instructions
1. **Install LM Studio**:
    - Download and install LM Studio from [official website or repository].
    - Run LM Studio locally and make sure it’s accessible at the URL specified in the `LM_STUDIO_URL` variable.
2. **Install DrawThings**:
    - Download and install DrawThings from [official website or repository].
    - Ensure that DrawThings is running locally and accessible via the URL specified in `DRAW_THINGS_URL`.
3. **Set Up Siri Shortcut for Background Removal**:
    - Create a Siri Shortcut named RemoveBackground on macOS. This shortcut should use an image editing application (like Preview or a custom tool) to remove backgrounds from images.
4. **Install Python Libraries**:
    - Install the required Python libraries by running the following commands:

    ```bash
    pip install requests python-pptx
    ```

5. **Run the Script**:
    - Select a PowerPoint file when prompted by the script.
    - The script will process each slide, generate descriptions and images, and add them to the PowerPoint file.

## Key Functions
1. **`extract_text_from_pptx(pptx_path)`**: Extracts text from each slide in the provided PowerPoint file.
2. **`description_prompt_text(prompt_text)`**: Sends slide text to LM Studio to generate an image description.
3. **`generate_image(prompt)`**: Sends the description to DrawThings to generate an image.
4. **`copy_image_to_clipboard(image_path)`**: Copies the generated image to the clipboard on macOS.
5. **`run_siri_shortcut(shortcut_name)`**: Executes a Siri Shortcut (used for background removal).
6. **`get_latest_downloaded_image()`**: Retrieves the latest downloaded image from the Downloads folder.
7. **`insert_image_to_slide(slide, image_path, x, y)`**: Inserts the image into the specified PowerPoint slide at the given coordinates.

## Troubleshooting
- **No description generated**: If LM Studio does not generate a description, the slide will be skipped.
- **No image generated**: If DrawThings fails to generate an image, the slide will be skipped.
- **No background-removed image found**: If the background removal shortcut fails, the script will use the original image.

## Example Usage
```bash
python generate_images.py
```

When prompted, select a PowerPoint file and let the script process each slide by generating images and converting text to speech.

## Notes
- Make sure that both LM Studio and DrawThings are running on your machine.
- The Siri Shortcut must be manually triggered the first time to give proper permissions.

This script is designed for macOS environments, utilizing system commands and local APIs to generate and manipulate images and text.
