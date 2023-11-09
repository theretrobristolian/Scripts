# Audio Normalization Script for Video Looping on Raspberry Pi

The Normalize-Audio.ps1 script is a versatile tool created to streamline the audio normalization process for video content. This script is particularly designed to work with two categories of video sources: music videos and trailers/advertisements. It processes these videos to ensure consistent and optimal audio levels. The resulting normalized videos are organized into separate folders. This organization is particularly useful for scenarios where you need distinct video content for specific applications, such as setting up a video looping system, such as on a Raspberry Pi.

The script essentially takes video files from these two source categories, applies audio normalization techniques to achieve consistent audio levels, and ensures that both audio sample rates and bitrates are preserved from the original sources. The output videos are neatly organized for easy access and use, making it a valuable tool for video content management and distribution.

## Prerequisites

Before using this script, ensure you have the following prerequisites in place:

### System Requirements

- **Operating System**: Windows 7 or later

### Software Dependencies

1. **7-Zip**: Install 7-Zip on your system. You can download it from [here](https://www.7-zip.org/download.html).

2. **FFmpeg**: Install FFmpeg on your system. You can download it from [here](https://ffmpeg.org/download.html#build-windows).

3. **PowerShell**: Have PowerShell installed and basic knowledge of its usage.

### Script and Folder Structure

- Download the script from this Git repository.
- The script can help you create a basic folder structure on the first run.

## Getting Started

1. Download the script from this repository.
2. Ensure you meet all the prerequisites listed above.
3. Run the script to do a first run and create the folder structure required.
4. using 7-zip extract the ffmpeg download and place the bin folder inin C:\Scripts\Apps\ and rename it to FFmpeg.

Now you're ready to use the `Normalize-Audio.ps1` script to prepare your videos for your Raspberry Pi video looping system. Enjoy a seamless and consistent audio experience!
