# ImageMapperLandingPageWebPart README

## Overview
This repository contains a SharePoint web part called **ImageMapperLandingPageWebPart**, which allows users to add image map areas to a webpage. The web part includes a customizable image with map areas that can be configured to perform various actions such as opening URLs in a new window or navigating to a specific link.

This web part integrates with SharePoint's property pane to allow easy configuration of image settings, such as the image's URL, size, position, and scale. Additionally, the user can add and delete map areas on the image.

## Features
- Customizable image URL, size, and position.
- Adjustable scale for the image.
- Add, delete, and configure map areas with support for different types (Rectangle or Path).
- Property Pane for configuring web part properties.

## Prerequisites
- A SharePoint Framework (SPFx) environment.
- Knowledge of React, TypeScript, and SharePoint development.
- SharePoint Online or On-Premises to deploy the web part.

## Getting Started

To get started with the Table of Contents WebPart, follow these steps:

1. Clone or download the repository to your local machine.
2. Navigate to the project folder in your terminal.
3. Install dependencies by running the following command:
   ```bash
   npm install

# Build and Serve the Web Part Locally

To build and serve your SharePoint Framework (SPFx) web part locally, follow these steps.

## Prerequisites

Ensure that you have the following installed:

- **Node.js**: The LTS version is recommended. You can download it from [nodejs.org](https://nodejs.org/).
  

  # Steps to Build and Serve the Web Part Locally

## 1. Navigate to Your Project Folder

Open a terminal and change to the folder where your SPFx project is located. You can do this by running the following command:

```bash
cd KWSPFX_ImageMapper/SPFx/

```

## 2. Install Project Dependencies

If you haven't already installed the required dependencies, run the following command to install them:

```bash
npm install

```

## 3. Build and Serve the Web Part

After installing the dependencies, you can build and serve your SPFx web part locally using the following command:

```bash
gulp serve

```
## 4. Access the Web Part

Once `gulp serve` is complete, a browser window should open automatically, displaying the local SharePoint workbench where you can test and preview your web part.

If the browser doesn't open automatically, you can access the workbench manually by visiting the following URL:

```text
https://001gcdev.sharePoint.com/sites/64617/temp/workbench.html

```

## 5. Test and Debug

You can now test your web part and make sure everything works as expected.

Any changes made to the code will be reflected immediately in the local workbench, allowing for live testing and debugging.

## Configuration Options

The following properties are available for configuration in the property pane:

### Image Settings
- **Image Url**  
  - Description: The URL of the image to be displayed.

- **Image Height**  
  - Description: The height of the image in pixels.

- **Image Width**  
  - Description: The width of the image in pixels.

- **Image Horizontal Position**  
  - Description: The horizontal position of the image.
  - Options:
    - Left
    - Center
    - Right

- **Image Vertical Position**  
  - Description: The vertical position of the image.
  - Options:
    - Top
    - Center
    - Bottom

- **Scale**  
  - Description: A slider to scale the image from 0 to 100.

### Map Areas
For each map area:

- **Map Area Type**  
  - Description: The type of the map area (either Rectangle or Path).

- **D (for Path)**  
  - Description: The path coordinates for the map area (used for `Path` type areas).

- **X**  
  - Description: The X coordinate for the map area (used for `Rectangle` type areas).

- **Y**  
  - Description: The Y coordinate for the map area (used for `Rectangle` type areas).

- **Width**  
  - Description: The width of the map area (used for `Rectangle` type areas).

- **Height**  
  - Description: The height of the map area (used for `Rectangle` type areas).

- **Url**  
  - Description: The URL to open when the map area is clicked.

- **Open in new window**  
  - Description: Whether the URL should open in a new window.
