# PPT Automate
Java library for automatizing template-based PPT production

## Introduction

PPT Automate creates PowerPoint presentations starting from a PPT template, data, and a set of action commands. Action commands can be either be written in Java or within a specific Groovy Script, i.e. to be dynamically stored and retrieved from a database.

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

TBD

## Available Action Commands

TBD

## Groovy Scripts

TBD
