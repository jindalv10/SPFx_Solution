# Table of Contents WebPart

This is a SharePoint WebPart that generates a Table of Contents based on the content of the page. It includes various customization options such as hiding the title, enabling sticky mode, and adjusting the list style.

## Features

- **Hide Title**: Option to hide the title of the web part.
- **Search Options**: Includes the ability to search for text, Markdown content, and collapsible content.
- **Heading Levels**: Configure which heading levels to display (Heading 1, 2, 3, 4).
- **Previous Page Link**: Add a previous page link with customization for positioning (above or below the content).
- **Sticky Mode**: Option to enable sticky behavior for the Table of Contents.
- **Mobile View**: Hide the web part in mobile views.
- **List Style**: Choose from different list styles (disc, circle, square, etc.).

## Prerequisites

Ensure you have the following tools installed:
- [Node.js](https://nodejs.org/) (for building and running the project)
- [SharePoint Framework (SPFx)](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/overview-client-side-web-parts)

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
cd path/to/your/project

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

  ## Properties Configuration

The following properties are available in the web part configuration:

| Property                         | Type    | Description                                                                 |
|-----------------------------------|---------|-----------------------------------------------------------------------------|
| `hideTitle`                       | boolean | Whether to hide the title of the Table of Contents.                         |
| `titleText`                       | string  | The title text to display for the Table of Contents.                        |
| `searchText`                      | boolean | Enable or disable searching for text.                                       |
| `searchMarkdown`                  | boolean | Enable or disable searching for Markdown content.                           |
| `searchCollapsible`               | boolean | Enable or disable searching for collapsible content.                        |
| `showHeading1`                    | boolean | Whether to show Heading 1 in the Table of Contents.                         |
| `showHeading2`                    | boolean | Whether to show Heading 2 in the Table of Contents.                         |
| `showHeading3`                    | boolean | Whether to show Heading 3 in the Table of Contents.                         |
| `showHeading4`                    | boolean | Whether to show Heading 4 in the Table of Contents.                         |
| `showPreviousPageLinkTitle`       | boolean | Whether to show the previous page link title.                               |
| `showPreviousPageLinkAbove`       | boolean | Whether to place the previous page link above the content.                  |
| `showPreviousPageLinkBelow`       | boolean | Whether to place the previous page link below the content.                  |
| `previousPageText`                | string  | The text for the previous page link.                                        |
| `enableStickyMode`                | boolean | Whether to enable sticky mode for the Table of Contents.                    |
| `hideInMobileView`                | boolean | Whether to hide the web part in mobile view.                                |
| `listStyle`                       | string  | Choose the list style (`default`, `disc`, `circle`, `square`, `none`).      |

## Property Pane Configuration

You can configure the properties of the web part via the property pane. The property pane includes options for the following:

- **Title**: Hide the title and set the title text.
- **Search Options**: Enable or disable search options for text, Markdown, and collapsible content.
- **Heading Levels**: Choose which heading levels to include (Heading 1-4).
- **Previous Page Link**: Configure the previous page link title and its positioning (above or below the content).
- **Sticky Mode**: Enable or disable sticky behavior for the web part.
- **Mobile View**: Hide the web part in mobile view if necessary.

## Usage

The web part will automatically render the Table of Contents based on the content available on the page.

You can toggle the visibility of headings, include a search function, or even enable a sticky mode to keep the Table of Contents visible as the user scrolls.

The configuration options in the property pane will allow you to customize the web part further based on your requirements.
