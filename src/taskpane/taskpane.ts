/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

    const dropZone = document.getElementById("dropZone") as HTMLDivElement;
    dropZone.addEventListener("dragover", (event) => {
      event.preventDefault();
    });

    dropZone.ondrop = fileDrop;
  }
});

export async function fileDrop(event: DragEvent) {
  event.preventDefault();
  event.stopPropagation();

  const files = event.dataTransfer.files;
  if (files.length > 0) {
    const file = files[0];

    const reader = new FileReader();

    reader.onload = (e2: ProgressEvent<FileReader>) => {
      const result = e2.target?.result;
      Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.load("height");
        range.load("width");
        range.load("left");
        range.load("top");
        await context.sync();

        const imageSize = await getImageSize(result.toString());

        let startIndex = result.toString().indexOf("base64,");
        let myBase64 = result.toString().substr(startIndex + 7);
        let sheet = context.workbook.worksheets.getActiveWorksheet();
        let image = sheet.shapes.addImage(myBase64);
        image.name = "Image";

        let ratio = range.width / imageSize.width;
        let height = imageSize.height * ratio;
        let width = range.width;

        if (height > range.height) {
          ratio = range.height / imageSize.height;
          height = range.height;
          width = imageSize.width * ratio;
        }

        image.height = height;
        image.width = width;

        image.left = range.left + (range.width - width) / 2;
        image.top = range.top + (range.height - height) / 2;

        return context.sync();
      }).catch(errorHandlerFunction);
    };

    // Read in the image file as a data URL.
    reader.readAsDataURL(file);
  }
}

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
function errorHandlerFunction(reason: any): PromiseLike<never> {
  throw new Error("Function not implemented." + reason.toString());
}

async function getImageSize(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const img = document.getElementById("preview") as HTMLImageElement;

    img.onload = () => {
      const size = {
        width: img.naturalWidth,
        height: img.naturalHeight,
      };

      resolve(size);
    };

    img.onerror = (error) => {
      reject(error);
    };

    img.src = dataUrl;
  });
}
