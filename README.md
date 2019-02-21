# PPT Automate
Java library for automatizing template-based PPT production

## Introduction
PPT Automate creates PowerPoint presentations starting from a PPT template, data, and a set of action commands. Action commands can be either be written in Java or within a specific Groovy Script, e.g. to be dynamically stored and retrieved from a database.

Processing a PPT can be as simple as

```
PptAutomate outputPpt = new PptAutomate(classloader.getResourceAsStream("template.pptx"));
outputPpt.executeGroovyScript(classloader.getResourceAsStream("script.groovy"));
outputPpt.finalizeAndWritePpt(new FileOutputStream(file));
```

and the Groovy Script can look something like

```
outputPpt
    .withAppendTemplateSlides([1, 2])
        .selectShapesMatchingRegex("TEXT.*")
            .setTextHtml("<ul><li>List item #1</li><li>List item #2</li></ul>")
```

Both PPT Template and the Groovy Script are expected as InputStream, and the processed PPT is written in an OutputStream.

## PPT Template
PPT Template is used by pptautomate as a base for creating all the slides of the presentation. When adding slides to the PPT to be generated, one or more slides of the PPT Template are indicated as indices. Such slides are then copied (appended) to the existing, previously created slides of the output PPT.

PPT Template must meet following criteria
* Must be in .pptx format
* Must contain at least one slide
* Must contain all slide templates that will be used in the output PPT
* Shapes should be named appropriately in order to be easily selected to apply Action Commands

A single template slide can be used multiple times throughout the output PPT with different content.
Please refer to the Available Action Commands section in order to know if specific care needs to be taken while preparing the PPT Template.

## Init of PPT Automate
PPT Automate can be instantiated with the following command

```
PptAutomate outputPpt = new PptAutomate(pptTemplateInputStream);
```

where pptTemplateInputStream is the PPT Template provided as InputStream.

Once initialized, PPT Automate will automatically setup an Output PPT, initially empty, to be processed with Action Commands provided via Java or via Groovy Script.

## Adding PPT Template slides to Output PPT
Template slides can be added to the Output PPT as follows:

```
outputPpt.withAppendTemplateSlides(templateSlideIndexes);
```

where templateSlideIndexes is the ArrayList of indexes of the template slides. Index numbering starts from 1 for the first slide of the Template PPT.

Chosen Template PPT slides are copied (appended) to the Output PPT and are automatically selected for Action Commands (see Selecting slides for Action Commands section). As many other methods of this library, withAppendTemplateSlides supports method chaining.

## Selecting slides and shapes for Action Commands
In order to perform Action Commands, both slides and shapes of the Output PPT need to be selected.

### Selecting slides
Slide selection can be done at anytime after at least one slide has already been copied from the Template PPT - otherwise an exception will be thrown. Also, slide indices must be within the range of [1, slideCount] where slideCount is the number of slides currently present into the Output PPT.

Selected slides indexes can be returned with
```
ArrayList<Integer> targetSlides = outputPpt.getTargetSlides();
```

#### Select all slides
```
outputPpt.selectAllOutputSlides();
```
#### Select a slide range
```
outputPpt.selectOutputSlides(idxStart, idxStop);
```
#### Select some slides
```
outputPpt.selectOutputSlides(idxArrayList);
```
#### Select one slide
```
outputPpt.selectOutputSlide(idx);
```

### Selecting shapes
Shape selection occurs within the scope of selected slides. Shapes can be selected either by name or by name pattern (regex). It is convenient to name the shapes of the Template PPT appropriately in order to be easily selected.

Selected shapes can be returned with
```
List<XSLFShape> targetSlides = outputPpt.getTargetShapes();
```

#### Select shapes by name
```
outputPpt.selectShapes(name);
```
#### Select shapes by name pattern (regex)
```
outputPpt.selectShapesMatchingRegex(regex);
```

## Action Commands
Action Commands are used to perform various actions on Output PPT shapes. Actions will not have any effect on Template PPT shapes. Action Commands methods must be called after selection of at least one shape - they will not have effect otherwise. Once called, Action Commands will be applied to all selected shapes sequentially.
### Fill Color
This Action Command is used to set the fill color of a shape.
```
outputPpt.fillColor(colorString);
outputPpt.fillColor(color);
```
* colorString is a String representation of a color - supported formats are rgb and hex (e.g. "rgb(0,0,0)" and "#000000")
* color is a java.awt.Color
### Move
This Action Command is used to move a shape within the containing slide.
```
outputPpt.move(position);
outputPpt.move(position, isRelative);
```
* position is a Position object contained in the pptautomate library that can be instantiated with "new Position(Double x, Double y)"
* isRelative is a boolean indicating if position indicates absolute coordinates or a relative displacement
### Resize
This Action Command is used to resize a shape.
```
outputPpt.resize(size);
outputPpt.resize(size, isRelative);
```
* size is a Size object contained in the pptautomate library that can be instantiated with "new Size(Double w, Double h)"
* isRelative is a boolean indicating if size indicates absolute measures or a relative displacement
### Delete
This Action Command is used to delete a shape.
```
outputPpt.delete(size);
```
### Replace with Image
This Action Command is used to replace a shape with an image. The shape will be deleted and replaced with a different shape with the same Z-Index but a different name - the original name of the shape will be lost. The image will be aligned with the top-left border of the original shape. If keepAspectRatio is true, the picture will fit the original shape rectangle by mantaining its proportions - default is true.
```
outputPpt.replaceWithImg(img, keepAspectRatio);
outputPpt.replaceWithImg(img);
```
* img is a Base64Image object contained in the pptautomate library that can be instantiated with "new Base64Image(byte[] data, PictureType type)"
* keepAspectRatio is a boolean indicating if the aspect ratio of the img is to be preserved
### Set HTML Text
This Action Command is used to fully replace the text of a shape. Keeping a minimal text into the shape of the Template PPT is strongly suggested, since pptautomate will take the font family, font size, and color of the first Text Run (snippet of text within the same paragraph with uniform properties), and alignment of the first Text Paragraph of the shape.

Plain text is accepted as well as some supported HTML tags:

| HTML Tag    | Effect                                                                |
| ----------- | --------------------------------------------------------------------- |
| \<strong\>  | Bold                                                                  |
| \<em\>      | Italic                                                                |
| \<u\>       | Underlined                                                            |
| \<ul\>      | Bullet Points (even nested for more levels)                           |
| \<ol\>      | Numbered list (even nested, but uses arabicPeriod schema by default)  |
| \<br/\>     | Line Break                                                            |
| \<p\>       | New Paragraph                                                         |

```
outputPpt.setTextHtml(string);
```
### Process Groovy GString
This Action Command allows to process shape text as Groovy strings (GString). All Text Runs of the shape are separately treated as GStrings and processed by the Groovy shell. It is important that expressions are contained into the same Text Run in order for the command to work and not result in exceptions. Please note that also underlined words for grammar corrections are processed as different Text Runs by PowerPoint - selecting "Ignore all" on the underlined word solves this issue. Keeping the processing at the Text Run level preserves the formatting (e.g. bold, font changes, bullet points) of the whole text of the shape.

All variables passed to the binding of the PptAutomate instance are also available to the Groovy shell used for this command - see "Passing variables to the Groovy shell" section below.
```
outputPpt.processText();
```
## Groovy Scripts
While Action Commands can be simply provided into the Java code, it can be useful to dynamically retrieve - or compose - a Groovy script. The script does not need to import PptAutomate, nor to return anything: PptAutomate adds these codelines for you and also passes the PptAutomate instance to the binding of the Groovy shell, the variable is called outputPpt.

Example Groovy script can look like:
```
outputPpt
	.withAppendTemplateSlides([1, 2])
		.selectShapesMatchingRegex("TEXT.*")
			.processText()
	.selectAllOutputSlides()
		.selectAllShapes()
		.selectShapes("LOGO")
		.replaceWithImg(img)
;
```

The Groovy script can be executed with:
```
outputPpt.executeGroovyScript(groovyScriptInputStream);
```
* groovyScriptInputStream is the Groovy script provided as InputStream
### Passing variables to the Groovy shell
Java objects can be made available to the Groovy shell by passing variables to the shell binding, e.g.:
```
outputPpt.getBinding().setVariable("num", 4);
```
## Saving the Output PPT
The Output PPT needs to be finalized before being written to an OutputStream, and a single method does this:
```
outputPpt.finalizeAndWritePpt(outputStream);
```
