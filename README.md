# PPT Automate
Java library for automatizing template-based PPT production

## Introduction
PPT Automate creates PowerPoint presentations starting from a PPT template, data, and a set of action commands. Action commands can be either be written in Java or within a specific Groovy Script, e.g. to be dynamically stored and retrieved from a database.

Processing a PPT can be as simple as

```
PptAutomate pptAutomate = new PptAutomate(classloader.getResourceAsStream("template.pptx"));
pptAutomate.executeGroovyScript(classloader.getResourceAsStream("script.groovy"));
pptAutomate.finalizeAndWritePpt(new FileOutputStream(file));
```

and the Groovy Script can look something like

```
pptAutomate
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
PptAutomate pptAutomate = new PptAutomate(pptTemplateInputStream);
```

where pptTemplateInputStream is the PPT Template provided as InputStream.

Once initialized, PPT Automate will automatically setup an Output PPT, initially empty, to be processed with Action Commands provided via Java or via Groovy Script.

## Adding PPT Template slides to Output PPT
Template slides can be added to the Output PPT as follows:

```
pptAutomate.withAppendTemplateSlides(templateSlideIndexes);
```

where templateSlideIndexes is the ArrayList of indexes of the template slides. Index numbering starts from 1 for the first slide of the Template PPT.

Chosen Template PPT slides are copied (appended) to the Output PPT and are automatically selected for Action Commands (see Selecting slides for Action Commands section). As many other methods of this library, withAppendTemplateSlides supports method chaining.

## Selecting slides and shapes for Action Commands
In order to perform Action Commands, both slides and shapes of the Output PPT need to be selected.

### Selecting slides
Slide selection can be done at anytime after at least one slide has already been copied from the Template PPT - otherwise an exception will be thrown. Also, slide indices must be within the range of [1, slideCount] where slideCount is the number of slides currently present into the Output PPT.

Selected slides indexes can be returned with
```
ArrayList<Integer> targetSlides = pptAutomate.getTargetSlides();
```

#### Select all slides
```
pptAutomate.selectAllOutputSlides();
```
#### Select a slide range
```
pptAutomate.selectOutputSlides(idxStart, idxStop);
```
#### Select some slides
```
pptAutomate.selectOutputSlides(idxArrayList);
```
#### Select one slide
```
pptAutomate.selectOutputSlide(idx);
```

### Selecting shapes
Shape selection occurs within the scope of selected slides. Shapes can be selected either by name or by name pattern (regex). It is convenient to name the shapes of the Template PPT appropriately in order to be easily selected.

Selected shapes can be returned with
```
List<XSLFShape> targetSlides = pptAutomate.getTargetShapes();
```

#### Select shapes by name
```
pptAutomate.selectShapes(name);
```
#### Select shapes by name pattern (regex)
```
pptAutomate.selectShapesMatchingRegex(regex);
```

## Action Commands
Action Commands are used to perform various actions on Output PPT shapes. Actions will not have any effect on Template PPT shapes. Action Commands methods must be called after selection of at least one shape - they will not have effect otherwise. Once called, Action Commands will be applied to all selected shapes sequentially.
### Fill Color
This Action Command is used to set the fill color of a shape.
```
pptAutomate.fillColor(colorString);
pptAutomate.fillColor(color);
```
* colorString is a String representation of a color - supported formats are rgb and hex (e.g. "rgb(0,0,0)" and "#000000")
* color is a java.awt.Color
### Move
This Action Command is used to move a shape within the containing slide.
```
pptAutomate.move(position);
pptAutomate.move(position, isRelative);
```
* position is a Position object contained in the pptautomate library that can be instantiated with "new Position(Double x, Double y)"
* isRelative is a boolean indicating if position indicates absolute coordinates or a relative displacement
### Resize
This Action Command is used to resize a shape.
```
pptAutomate.resize(size);
pptAutomate.resize(size, isRelative);
```
* size is a Size object contained in the pptautomate library that can be instantiated with "new Size(Double w, Double h)"
* isRelative is a boolean indicating if size indicates absolute measures or a relative displacement
### Delete
### Replace with Image
### Set HTML Text
### Process Groovy GString


## Groovy Scripts
TBD
### Binding and Variables

## Saving the Output PPT
TBD
