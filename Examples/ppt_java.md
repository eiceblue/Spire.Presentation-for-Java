# Spire.Presentation Hello World
## Create a simple PowerPoint presentation with Hello World text
```java
//Create PPT document
Presentation presentation = new Presentation();

//Add new shape to PPT document
Rectangle rec = new Rectangle((int)presentation.getSlideSize().getSize().getWidth() / 2 - 250, 80, 500, 150);
IAutoShape shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, rec);

shape.getShapeStyle().getLineColor().setColor(Color.white);
shape.getFill().setFillType(FillFormatType.NONE);
//Add text to shape
shape.appendTextFrame("Hello World!");

//Set the font and fill style of text
PortionEx textRange = shape.getTextFrame().getTextRange();
textRange.getFill().setFillType(FillFormatType.SOLID);
textRange.getFill().getSolidColor().setColor(Color.cyan);
textRange.setFontHeight(66);
textRange.setLatinFont(new TextFont("Lucida Sans Unicode"));
```

---

# Spire.Presentation Math Equation
## Add mathematical equations in LaTeX format to paragraphs in a presentation
```java
// Create a presentation
Presentation ppt = new Presentation();

// Append shape
IAutoShape shape=ppt.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE,new Rectangle2D.Float(30,100,400,200));
// Clear shape
shape.getTextFrame().getParagraphs().clear();

// Append paragraph
ParagraphEx p=new ParagraphEx();
shape.getTextFrame().getParagraphs().append(p);

// Append text and latex code
PortionEx portionEx=new PortionEx("Test");
p.getTextRanges().append(portionEx);
// Add LaTeX math equation
p.appendFromLatexMathCode("latex_math_code");
PortionEx portionEx2=new PortionEx("Hello");
p.getTextRanges().append(portionEx2);
```

---

# Spire Presentation Paragraph Management
## Add and format paragraph in presentation
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Append a new shape
IAutoShape shape = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle(50, 70, 620, 150));
shape.getFill().setFillType(FillFormatType.NONE);
shape.getShapeStyle().getLineColor().setColor(Color.white);

//Set the alignment of paragraph
shape.getTextFrame().getParagraphs().get(0).setAlignment(TextAlignmentType.LEFT);
//Set the indent of paragraph
shape.getTextFrame().getParagraphs().get(0).setIndent(50);
//Set the linespacing of paragraph
shape.getTextFrame().getParagraphs().get(0).setLineSpacing(150);
//Set the text of paragraph
shape.getTextFrame().setText("This powerful component suite contains the most up-to-date versions of all .NET WPF Silverlight components offered by E-iceblue.");

//Set the Font
shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0).setLatinFont(new TextFont("Arial Rounded MT Bold"));
shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0).getFill().setFillType(FillFormatType.SOLID);
shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0).getFill().getSolidColor().setColor(Color.BLACK);
```

---

# Spire.Presentation Text Alignment
## Set text alignment for paragraphs in a PowerPoint slide
```java
//Get the related shape and set the text alignment
IAutoShape shape = (IAutoShape) presentation.getSlides().get(0).getShapes().get(1);
shape.getTextFrame().getParagraphs().get(0).setAlignment(TextAlignmentType.LEFT);
shape.getTextFrame().getParagraphs().get(1).setAlignment(TextAlignmentType.CENTER);
shape.getTextFrame().getParagraphs().get(2).setAlignment(TextAlignmentType.RIGHT);
shape.getTextFrame().getParagraphs().get(3).setAlignment(TextAlignmentType.JUSTIFY);
shape.getTextFrame().getParagraphs().get(4).setAlignment(TextAlignmentType.NONE);
```

---

# Spire.Presentation HTML Appending
## Append HTML content to PowerPoint shapes
```java
//Add a shape
IAutoShape shape = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE,
        new Rectangle(150, 100, 200, 200));

//Clear default paragraphs
shape.getTextFrame().getParagraphs().clear();

String code = "<html><body><p>This is a paragraph</p></body></html>";

//Append HTML, and generate a paragraph with default style in PPT document.
shape.getTextFrame().getParagraphs().addFromHtml(code);

String codeColor = "<html><body><p style=\" color:black \">This is a paragraph</p></body></html>";
//Append HTML with black setting
shape.getTextFrame().getParagraphs().addFromHtml(codeColor);

//Add another shape
IAutoShape shape1 = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle(350, 100, 200, 200));

//Clear default paragraph
shape1.getTextFrame().getParagraphs().clear();

//Change the fill format of shape
shape1.getFill().setFillType(FillFormatType.SOLID);
shape1.getFill().getSolidColor().setColor(Color.white);

//Append HTML
shape1.getTextFrame().getParagraphs().addFromHtml(code);
ParagraphEx par = shape1.getTextFrame().getParagraphs().get(0);

//Change the fill color for paragraph
PortionEx textRange = shape1.getTextFrame().getTextRange();
textRange.getFill().setFillType(FillFormatType.SOLID);
textRange.getFill().getSolidColor().setColor(Color.black);
```

---

# Spire Presentation AutoFit Text or Shape
## Demonstrates how to auto-fit text within shapes or resize shapes to fit text
```java
//Set the AutofitType property to Shape
IAutoShape textShape2 = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle(150, 100, 150, 80));
textShape2.getTextFrame().setText("Resize shape to fit text.");
textShape2.getTextFrame().setAutofitType(TextAutofitType.SHAPE);

//Set the AutofitType property to Normal
IAutoShape textShape1 = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle(400, 100, 150, 80));
textShape1.getTextFrame().setText("Shrink text to fit shape. Shrink text to fit shape. Shrink text to fit shape. Shrink text to fit shape.");
textShape1.getTextFrame().setAutofitType(TextAutofitType.NORMAL);
```

---

# Spire Presentation Shape Borders and Shading
## Set border, gradient fill, and shadow effects for shapes in PowerPoint presentations
```java
IAutoShape shape = (IAutoShape)presentation.getSlides().get(0).getShapes().get(0);

//Set line color and width of the border
shape.getLine().setFillType(FillFormatType.SOLID);
shape.getLine().setWidth(3);
shape.getLine().getSolidFillColor().setColor(Color.yellow);

//Set the gradient fill color of shape
shape.getFill().setFillType(FillFormatType.GRADIENT);
shape.getFill().getGradient().setGradientShape(GradientShapeType.LINEAR);
shape.getFill().getGradient().getGradientStops().append(1f, KnownColors.LIGHT_BLUE);
shape.getFill().getGradient().getGradientStops().append(0, KnownColors.LIGHT_SKY_BLUE);

//Set the shadow for the shape
OuterShadowEffect shadow = new OuterShadowEffect();
shadow.setBlurRadius(20);
shadow.setDirection (30);
shadow.setDistance (8);
shadow.getColorFormat().setColor( Color.green);
shape.getEffectDag().setOuterShadowEffect(shadow);
```

---

# Spire Presentation Bullets
## Add numbered bullets with Roman numeral style to paragraphs in a presentation shape
```java
// Get shape from presentation
IAutoShape shape = (IAutoShape) presentation.getSlides().get(0).getShapes().get(1);

// Add bullets to paragraphs
for (Object para : shape.getTextFrame().getParagraphs()) {
    ParagraphEx paragraph = (ParagraphEx) para;
    paragraph.setBulletType(TextBulletType.NUMBERED);
    paragraph.setBulletStyle(NumberedBulletStyle.BULLET_ROMAN_LC_PERIOD);
}
```

---

# Change Text Style in Presentation
## Modify text properties like color, font, size and underline in presentation paragraphs
```java
// Get shape from presentation
IAutoShape shape = (IAutoShape) presentation.getSlides().get(0).getShapes().get(0);
ParagraphCollection paras = shape.getTextFrame().getParagraphs();

// Set the style for the text content in the first paragraph
ParagraphEx para1 = paras.get(0);
int max = para1.getTextRanges().getCount();
for (int i = 0; i < max; i++) {
    para1.getTextRanges().get(i).getFill().setFillType(FillFormatType.SOLID);
    para1.getTextRanges().get(i).getFill().getSolidColor().setColor(Color.GREEN);
    para1.getTextRanges().get(i).setLatinFont(new TextFont("Lucida Sans Unicode"));
    para1.getTextRanges().get(i).setFontHeight(14);
}

// Set the style for the text content in the third paragraph
ParagraphEx para3 = paras.get(2);
max = para3.getTextRanges().getCount();
for (int i = 0; i < max; i++) {
    para3.getTextRanges().get(i).getFill().setFillType(FillFormatType.SOLID);
    para3.getTextRanges().get(i).getFill().getSolidColor().setColor(Color.blue);
    para3.getTextRanges().get(i).setLatinFont(new TextFont("Calibri"));
    para3.getTextRanges().get(i).setFontHeight(16);
    para3.getTextRanges().get(i).setTextUnderlineType(TextUnderlineType.DASHED);
}
```

---

# Spire.Presentation Text Copying
## Copy paragraph text from one PowerPoint presentation to another
```java
// Get the text from the first shape on the first slide
IShape sourceshp = ppt1.getSlides().get(0).getShapes().get(0);
String text1 = ((IAutoShape) sourceshp).getTextFrame().getText();

// Get the first shape on the first slide from the target file
IShape destshp = ppt2.getSlides().get(0).getShapes().get(0);

// Add the text to the target file
String text2 = ((IAutoShape) destshp).getTextFrame().getText();
((IAutoShape) destshp).getTextFrame().setText(text2 + "\n\n" + text1);
```

---

# Custom Bullet Numbering in PowerPoint
## Set custom numbered bullets for paragraphs in a presentation
```java
//Get the first slide
ISlide slide = presentation.getSlides().get(0);

//Access the first placeholder in the slide and typecasting it as AutoShape
ITextFrameProperties tf1 = ((IAutoShape) slide.getShapes().get(1)).getTextFrame();

//Access the first Paragraph and set bullet style
ParagraphEx para = tf1.getParagraphs().get(0);
para.setDepth((short) 0);
para.setBulletType(TextBulletType.NUMBERED);
para.setBulletStyle(NumberedBulletStyle.BULLET_ARABIC_PERIOD);
para.setBulletNumber((short) 2);

//Access the second Paragraph and set bullet style
para = tf1.getParagraphs().get(1);
para.setDepth((short) 0);
para.setBulletType(TextBulletType.NUMBERED);
para.setBulletStyle(NumberedBulletStyle.BULLET_ARABIC_PERIOD);
para.setBulletNumber((short) 4);

//Access the third Paragraph and set bullet style
para = tf1.getParagraphs().get(2);
para.setDepth((short) 0);
para.setBulletType(TextBulletType.NUMBERED);
para.setBulletStyle(NumberedBulletStyle.BULLET_ARABIC_PERIOD);
para.setBulletNumber((short) 6);

//Access the fourth Paragraph and set bullet style
para = tf1.getParagraphs().get(3);
para.setDepth((short) 0);
para.setBulletType(TextBulletType.NUMBERED);
para.setBulletStyle(NumberedBulletStyle.BULLET_ARABIC_PERIOD);
para.setBulletNumber((short) 7);
```

---

# Extract Text from PowerPoint Presentation
## This code demonstrates how to extract text from slides in a PowerPoint presentation
```java
// Create a PPT document
Presentation presentation = new Presentation();

// StringBuilder to store extracted text
StringBuilder buffer = new StringBuilder();

// Extract text from each slide
for (Object slide : presentation.getSlides()) {
    for (Object shape : ((ISlide) slide).getShapes()) {
        if (shape instanceof IAutoShape) {
            for (Object tp : ((IAutoShape) shape).getTextFrame().getParagraphs()) {
                buffer.append(((ParagraphEx) tp).getText());
            }
        }
    }
}
```

---

# Spire.Presentation Find and Format Text
## Find the first occurrence of specific text in a PowerPoint presentation and apply formatting
```java
// Find the specified text in the first shape
String text = "Spire.Presentation";
PortionEx textRange=ppt.getSlides().get(0).getFindFirstTextAsRange(text);

// Set text format
textRange.getFill().setFillType(FillFormatType.SOLID);
textRange.getFill().getSolidColor().setColor(Color.red);
textRange.setFontHeight(28);
textRange.setLatinFont(new TextFont("Arial"));
textRange.isItalic(TriState.TRUE);
textRange.setTextUnderlineType(TextUnderlineType.DOUBLE);
```

---

# Spire.Presentation Text Frame Data Extraction
## Extract text frame effective data from PowerPoint slides
```java
//Create a PPT document
Presentation presentation = new Presentation();

//Get the first slide
ISlide slide = presentation.getSlides().get(0);
//Get a shape 
IAutoShape shape = (IAutoShape) presentation.getSlides().get(0).getShapes().get(0);

ITextFrameProperties textFrameFormat = shape.getTextFrame();
```

---

# Spire.Presentation Text Style Extraction
## Extract text style effective data from PowerPoint presentation shapes
```java
//Create a PPT document
Presentation presentation = new Presentation();

//Get the first slide
ISlide slide = presentation.getSlides().get(0);
//Get a shape 
IAutoShape shape = (IAutoShape) presentation.getSlides().get(0).getShapes().get(0);

StringBuilder str = new StringBuilder();
for (int p = 0; p < shape.getTextFrame().getParagraphs().getCount(); p++) {
    ParagraphEx paragraph = shape.getTextFrame().getParagraphs().get(p);
    str.append("Text style for Paragraph " + p + " :" + "\r\n");
    //Get the paragraph style
    str.append(" Indent: " + paragraph.getIndent() + "\r\n");
    str.append(" Alignment: " + paragraph.getAlignment() + "\r\n");
    str.append(" Font alignment: " + paragraph.getFontAlignment() + "\r\n");
    str.append(" Hanging punctuation: " + paragraph.getHangingPunctuation() + "\r\n");
    str.append(" Line spacing: " + paragraph.getLineSpacing() + "\r\n");
    str.append(" Space before: " + paragraph.getSpaceBefore() + "\r\n");
    str.append(" Space after: " + paragraph.getSpaceAfter() + "\r\n");
    for (int r = 0; r < paragraph.getTextRanges().getCount(); r++) {
        PortionEx textRange = paragraph.getTextRanges().get(r);
        str.append("  Text style for Paragraph " + p + " TextRange " + r + " :" + "\r\n");
        //Get the text range style
        str.append("    Font height: " + textRange.getFontHeight() + "\r\n");
        str.append("    Language: " + textRange.getLanguage() + "\r\n");
        str.append("    Font: " + textRange.getLatinFont().getFontName() + "\r\n");
    }
}
```

---

# Spire Presentation Text Highlighting
## Highlight specified text in a presentation with custom options
```java
//Get the specified shape
IAutoShape shape = (IAutoShape)presentation.getSlides().get(0).getShapes().get(1);

//Set highlight options
TextHighLightingOptions options = new TextHighLightingOptions();
options.setWholeWordsOnly(true);
options.setCaseSensitive(true);

//Highlight text
shape.getTextFrame().highLightText("Spire", Color.yellow, options);
```

---

# Spire Presentation Paragraph Indentation
## Set paragraph indentation in PowerPoint presentation
```java
// Get the shape from the first slide
IAutoShape shape = (IAutoShape) presentation.getSlides().get(0).getShapes().get(0);

// Set the paragraph style for first paragraph
shape.getTextFrame().getParagraphs().get(0).setIndent(20);
shape.getTextFrame().getParagraphs().get(0).setLeftMargin(10);
shape.getTextFrame().getParagraphs().get(0).setSpaceAfter(10);

// Set the paragraph style of the third paragraph
shape.getTextFrame().getParagraphs().get(2).setIndent(-100);
shape.getTextFrame().getParagraphs().get(2).setLeftMargin(40);
shape.getTextFrame().getParagraphs().get(2).setSpaceBefore(0);
shape.getTextFrame().getParagraphs().get(2).setSpaceAfter(0);
```

---

# Spire Presentation HTML with Image
## Insert HTML content with image into a presentation slide
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();
ShapeList shapes = ppt.getSlides().get(0).getShapes();

shapes.addFromHtml("<html><div><p>First paragraph</p><p><img src='data\\logo.png'/></p><p>Second paragraph </p></html>");
```

---

# Spire.Presentation Line Spacing
## Set line spacing for text in PowerPoint slides
```java
//Add a shape
IAutoShape shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE,
        new Rectangle(50, 100, (int) presentation.getSlideSize().getSize().getWidth() - 100, 300));
shape.getShapeStyle().getLineColor().setColor(Color.WHITE);
shape.getFill().setFillType(FillFormatType.NONE);
shape.getTextFrame().getParagraphs().clear();

//Add text
shape.appendTextFrame("Spire.Presentation for Java is a professional PowerPoint API that enables developers"
        + " to create, read, write, convert and save PowerPoint documents in Java Applications. As an independent"
        + " Java library, Spire.Presentation doesn't need Microsoft PowerPoint to be installed on system.");
//Set font and color of text
PortionEx textRange = shape.getTextFrame().getTextRange();
textRange.getFill().setFillType(FillFormatType.SOLID);
shape.getShapeStyle().getLineColor().setColor(Color.BLUE);
textRange.setFontHeight(20);
textRange.setLatinFont(new TextFont("Lucida Sans Unicode"));

//Set properties of paragraph
shape.getTextFrame().getParagraphs().get(0).setSpaceBefore(100);
shape.getTextFrame().getParagraphs().get(0).setSpaceAfter(100);
shape.getTextFrame().getParagraphs().get(0).setLineSpacing(150);
```

---

# Spire Presentation Mixed Font Styles
## Apply different font styles to text portions in a presentation
```java
//Get the shape from the first slide
IAutoShape shape = (IAutoShape) presentation.getSlides().get(0).getShapes().get(0);

//Get the paragraph from TextRange
ParagraphEx tp = shape.getTextFrame().getTextRange().getParagraph();

//Append normal text that is in front of 'bold' to the paragraph
PortionEx tr = new PortionEx("\r\nHere is testing text. Only a few words are in ");
tp.getTextRanges().append(tr);
//Set font style of the text 'bold' as bold
tr = new PortionEx("bold");
tr.isBold(TriState.TRUE);
tp.getTextRanges().append(tr);

//Append normal text that is in front of 'red' to the paragraph
tr = new PortionEx(", some are in ");
tp.getTextRanges().append(tr);
//Set the color of the text 'red' as red
tr = new PortionEx("red");
tr.getFill().setFillType(FillFormatType.SOLID);
tr.getFormat().getFill().getSolidColor().setColor(Color.RED);
tp.getTextRanges().append(tr);

//Append normal text that is in front of 'underlined' to the paragraph
tr = new PortionEx(" color, some are ");
tp.getTextRanges().append(tr);
//Underline the text 'undelined'
tr = new PortionEx("underlined");
tr.setTextUnderlineType(TextUnderlineType.SINGLE);
tp.getTextRanges().append(tr);

//Append normal text that is in front of 'bigger font size' to the paragraph
tr = new PortionEx(", and some are in ");
tp.getTextRanges().append(tr);
//Set a large font for the text 'bigger font size'
tr = new PortionEx("bigger font size");
tr.setFontHeight(35);
tp.getTextRanges().append(tr);

//Append other normal text
tr = new PortionEx(".");
tp.getTextRanges().append(tr);
```

---

# Spire Presentation Multiple Level Bullets
## Create multiple level bullets in a PowerPoint presentation
```java
//Access the first placeholder in the slide and typecasting it as AutoShape
ITextFrameProperties tf1 = ((IAutoShape) slide.getShapes().get(1)).getTextFrame();

//Access the first Paragraph and set bullet style
ParagraphEx para = tf1.getParagraphs().get(0);
para.setBulletType(TextBulletType.SYMBOL);
para.setBulletChar(Convert.toChar(8226));
para.setDepth((short) 0);

//Access the second Paragraph and set bullet style
para = tf1.getParagraphs().get(1);
para.setBulletType(TextBulletType.SYMBOL);
para.setBulletChar('-');
para.setDepth((short) 1);

//Access the third Paragraph and set bullet style
para = tf1.getParagraphs().get(2);
para.setBulletType(TextBulletType.SYMBOL);
para.setBulletChar(Convert.toChar(8226));
para.setDepth((short) 2);

//Access the fourth Paragraph and set bullet style
para = tf1.getParagraphs().get(3);
para.setBulletType(TextBulletType.SYMBOL);
para.setBulletChar('-');
para.setDepth((short) 3);
```

---

# Spire.Presentation Multiple Paragraphs
## Create and format multiple paragraphs in a PowerPoint presentation
```java
//Create a PPT document
Presentation presentation = new Presentation();

//Access the first slide
ISlide slide = presentation.getSlides().get(0);

// Add an AutoShape of rectangle type
Rectangle rec = new Rectangle((int) presentation.getSlideSize().getSize().getWidth() / 2 - 250, 150, 500, 150);
IAutoShape shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, rec);

// Access TextFrame of the AutoShape
ITextFrameProperties tf = shape.getTextFrame();

// Create Paragraphs and PortionExs with different text formats
ParagraphEx para0 = tf.getParagraphs().get(0);
PortionEx PortionEx1 = new PortionEx();
PortionEx PortionEx2 = new PortionEx();
para0.getTextRanges().append(PortionEx1);
para0.getTextRanges().append(PortionEx2);

ParagraphEx para1 = new ParagraphEx();
tf.getParagraphs().append(para1);
PortionEx PortionEx11 = new PortionEx();
PortionEx PortionEx12 = new PortionEx();
PortionEx PortionEx13 = new PortionEx();
para1.getTextRanges().append(PortionEx11);
para1.getTextRanges().append(PortionEx12);
para1.getTextRanges().append(PortionEx13);

ParagraphEx para2 = new ParagraphEx();
tf.getParagraphs().append(para2);
PortionEx PortionEx21 = new PortionEx();
PortionEx PortionEx22 = new PortionEx();
PortionEx PortionEx23 = new PortionEx();
para2.getTextRanges().append(PortionEx21);
para2.getTextRanges().append(PortionEx22);
para2.getTextRanges().append(PortionEx23);

for (int i = 0; i < 3; i++)
    for (int j = 0; j < 3; j++) {
        tf.getParagraphs().get(i).getTextRanges().get(j).setText("TextRange " + j);
        if (j == 0) {
            tf.getParagraphs().get(i).getTextRanges().get(j).getFill().setFillType(FillFormatType.SOLID);
            tf.getParagraphs().get(i).getTextRanges().get(j).getFill().getSolidColor().setColor(Color.cyan);
            tf.getParagraphs().get(i).getTextRanges().get(j).getFormat().isBold(TriState.TRUE);
            tf.getParagraphs().get(i).getTextRanges().get(j).setFontHeight(15);
        } else if (j == 1) {
            tf.getParagraphs().get(i).getTextRanges().get(j).getFill().setFillType(FillFormatType.SOLID);
            tf.getParagraphs().get(i).getTextRanges().get(j).getFill().getSolidColor().setColor(Color.BLUE);
            tf.getParagraphs().get(i).getTextRanges().get(j).getFormat().isBold(TriState.TRUE);
            tf.getParagraphs().get(i).getTextRanges().get(j).setFontHeight(18);
        }
    }
```

---

# Picture Custom Bullet Style
## Set custom picture as bullet style for paragraphs in presentation
```java
//Get the second shape on the first slide
IAutoShape shape = (IAutoShape) ppt.getSlides().get(0).getShapes().get(1);

//Traverse through the paragraphs in the shape
for (Object para : shape.getTextFrame().getParagraphs()) {
    //Set the bullet style of paragraph as picture
    ParagraphEx paragraph = (ParagraphEx) para;
    paragraph.setBulletType(TextBulletType.PICTURE);
    //Load a picture
    BufferedImage bulletPicture = ImageIO.read(new File("data/icon.png"));
    //Add the picture as the bullet style of paragraph
    paragraph.getBulletPicture().setEmbedImage(ppt.getImages().append(bulletPicture));
}
```

---

# Remove TextBox from PowerPoint Slide
## This code demonstrates how to remove text boxes from a PowerPoint presentation slide
```java
//Get the first slide
ISlide slide = ppt.getSlides().get(0);

//Traverse all the shapes in slide
for (int i = 0; i < slide.getShapes().getCount(); i++) {
    if(slide.getShapes().get(i).getName().contains("TextBox")){
        slide.getShapes().removeAt(i);
        i--;
    }
}
```

---

# Spire.Presentation Text Replacement and Formatting
## Replace text with specific formatting in PowerPoint presentation
```java
// Create a new object to store the default text range formatting properties.
PortionFormatEx format = new PortionFormatEx();

// Set the IsBold property of the text range formatting to true, making the text bold.
format.isBold(TriState.TRUE);

// Set the FillType property of the text range fill to Solid, indicating a solid fill color.
format.getFill().setFillType(FillFormatType.SOLID);

// Set the Color property of the solid fill color to red.
format.getFill().getSolidColor().setColor(Color.red);

// Set the FontHeight property of the text range formatting to 45, indicating the font size.
format.setFontHeight(45);

// Replace all occurrences of the text "Spire.Presentation" with "Spire.PPT" and apply the specified formatting.
ppt.ReplaceAndFormatText("Spire.Presentation", "Spire.PPT", format);
```

---

# Spire.Presentation Text Replacement
## Replace text in PowerPoint presentation slides
```java
public static void replaceTags(ISlide pSlide, Map<String, String> tagValues) {
    for (int i = 0; i < pSlide.getShapes().getCount(); i++) {
        IShape curShape = pSlide.getShapes().get(i);
        if (curShape instanceof IAutoShape) {
            for (int j = 0; j < ((IAutoShape) curShape).getTextFrame().getParagraphs().getCount(); j++) {
                ParagraphEx tp = ((IAutoShape) curShape).getTextFrame().getParagraphs().get(j);
                for (Map.Entry<String, String> entry : tagValues.entrySet()) {
                    String mapKey = entry.getKey();
                    String mapValue = entry.getValue();
                    if (tp.getText().contains(mapKey)) {
                        tp.setText(tp.getText().replace(mapKey, mapValue));
                    }
                }
            }
        }
    }
}
```

---

# Spire Presentation Text Replacement
## Replace text in presentation while retaining style
```java
//Create PPT document and load file
Presentation presentation = new Presentation();

//Replace first occurrence of text on first slide
presentation.getSlides().get(0).replaceFirstText("use", "test", true);

//Replace all occurrences of text on second slide
presentation.getSlides().get(1).replaceAllText("Spire", "new spire", true);
```

---

# Spire Presentation Text Rotation
## Rotate text in a presentation shape
```java
// Get the first slide
ISlide slide = presentation.getSlides().get(0);
// Get a shape
IAutoShape shape = (IAutoShape) slide.getShapes().get(0);

// Set text rotation to 270 degrees vertical
shape.getTextFrame().setVerticalTextType(VerticalTextType.VERTICAL_270);
```

---

# Spire.Presentation 3D Text Effects
## Set 3D effects for text in a presentation slide
```java
//Add text to the shape
shape.appendTextFrame("This demo shows how to add 3D effect text to Presentation slide");

//Set text properties
PortionEx textRange = shape.getTextFrame().getTextRange();
textRange.getFill().setFillType(FillFormatType.SOLID);
textRange.getFill().getSolidColor().setColor(Color.BLUE);
textRange.setFontHeight(40);
textRange.setLatinFont(new TextFont("Arial"));

//Set 3D effect for text
shape.getTextFrame().getTextThreeD().getShapeThreeD().setPresetMaterial(PresetMaterialType.MATTE);
shape.getTextFrame().getTextThreeD().getLightRig().setPresetType(PresetLightRigType.SUNRISE);
shape.getTextFrame().getTextThreeD().getShapeThreeD().getTopBevel().setPresetType(BevelPresetType.CIRCLE);
shape.getTextFrame().getTextThreeD().getShapeThreeD().getContourColor().setColor(Color.BLUE);
shape.getTextFrame().getTextThreeD().getShapeThreeD().setContourWidth(3);
```

---

# Spire.Presentation Text Frame Anchor
## Set the anchor type of a text frame in a presentation
```java
//Get the first slide
ISlide slide = presentation.getSlides().get(0);
//Get a shape
IAutoShape shape = (IAutoShape) slide.getShapes().get(0);
shape.getTextFrame().setAnchoringType(TextAnchorType.BOTTOM);
```

---

# Spire Presentation Text Frame Column Count
## Set column count for text frames in presentation slides
```java
//Get the first shape in first slide and set column count of text
IAutoShape shape1 = (IAutoShape)ppt.getSlides().get(0).getShapes().get(0);
shape1.getTextFrame().setColumnCount(2);

//Get the second shape in second slide and set column count of text
IAutoShape shape2 = (IAutoShape)ppt.getSlides().get(1).getShapes().get(0);
shape2.getTextFrame().setColumnCount(3);
```

---

# Spire.Presentation Custom Fonts
## Set custom fonts in PowerPoint presentation
```java
//Create PPT document and add text
Presentation presentation = new Presentation();
IAutoShape shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle(50, 50, 200, 100));
shape.appendTextFrame("Hello World!");

//Set the folder of custom fonts
presentation.setCustomFontsFolder("/customFonts_folder");

//Set custom font for text
PortionEx textRange = shape.getTextFrame().getTextRange();
textRange.setLatinFont(new TextFont("Open Sans"));
```

---

# Spire Presentation Paragraph Font Styling
## Demonstrates how to set font properties for paragraphs in a presentation
```java
// Access the first and second placeholder in the slide and typecasting it as AutoShape
ITextFrameProperties tf1 = ((IAutoShape) slide.getShapes().get(0)).getTextFrame();
ITextFrameProperties tf2 = ((IAutoShape) slide.getShapes().get(1)).getTextFrame();

// Access the first Paragraph
ParagraphEx para1 = tf1.getParagraphs().get(0);
ParagraphEx para2 = tf2.getParagraphs().get(0);

// Justify the paragraph
para2.setAlignment(TextAlignmentType.JUSTIFY);

// Access the first text range
PortionEx textRange1 = para1.getFirstTextRange();
PortionEx textRange2 = para2.getFirstTextRange();

// Define new fonts
TextFont fd1 = new TextFont("Elephant");
TextFont fd2 = new TextFont("Castellar");

// Assign new fonts to text range
textRange1.setLatinFont(fd1);
textRange2.setLatinFont(fd2);

// Set font to Bold
textRange1.getFormat().isBold(TriState.TRUE);
textRange2.getFormat().isBold(TriState.FALSE);

// Set font to Italic
textRange1.getFormat().isItalic(TriState.FALSE);
textRange2.getFormat().isItalic(TriState.TRUE);

// Set font color
textRange1.getFill().setFillType(FillFormatType.SOLID);
textRange1.getFill().getSolidColor().setColor(Color.blue);
textRange1.getFill().setFillType(FillFormatType.SOLID);
textRange2.getFill().getSolidColor().setColor(Color.ORANGE);
```

---

# Spire Presentation Text Shadow Effect
## Set shadow effect for text in presentation slides
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Get reference of the slide
ISlide slide = ppt.getSlides().get(0);

//Add a new rectangle shape to the first slide
IAutoShape shape = slide.getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle(120, 100, 450, 200));
shape.getFill().setFillType(FillFormatType.NONE);

//Add the text to the shape and set the font for the text
shape.appendTextFrame("Text shading on slides");
shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0).setLatinFont(new TextFont("Arial Black"));
shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0).setFontHeight(21);
shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0).getFill().setFillType(FillFormatType.SOLID);
shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0).getFill().getSolidColor().setColor(Color.BLACK);

//Add outer shadow and set all necessary parameters
OuterShadowEffect Shadow = new OuterShadowEffect();

Shadow.setBlurRadius(0);
Shadow.setDirection(50);
Shadow.setDistance(10);
Shadow.getColorFormat().setColor(new Color(0xAD,0xD8,0xE6));

shape.getTextFrame().getTextRange().getEffectDag().setOuterShadowEffect(Shadow);
```

---

# Spire.Presentation Text Direction
## Set vertical text direction in presentation slides
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Append a shape with text to the first slide
IAutoShape textboxShape = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle(250, 70, 100, 400));
textboxShape.getShapeStyle().getLineColor().setColor(Color.white);
textboxShape.getFill().setFillType(FillFormatType.SOLID);
textboxShape.getFill().getSolidColor().setColor(Color.cyan);
textboxShape.getTextFrame().setText("You Are Welcome Here");
//Set the text direction to vertical
textboxShape.getTextFrame().setVerticalTextType(VerticalTextType.VERTICAL);

//Append another shape with text to the slide
textboxShape = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle(350, 70, 100, 400));
textboxShape.getShapeStyle().getLineColor().setColor(Color.white);
textboxShape.getFill().setFillType(FillFormatType.SOLID);
textboxShape.getFill().getSolidColor().setColor(Color.lightGray);
//Append some asian characters
textboxShape.getTextFrame().setText("Welcome");
//Set the VerticalTextType as EastAsianVertical to avoid rotating text 90 degrees
textboxShape.getTextFrame().setVerticalTextType(VerticalTextType.EAST_ASIAN_VERTICAL);
```

---

# Spire.Presentation Text Font Properties
## Set various font properties for text in a presentation
```java
//Get text range from shape
PortionEx textRange = shape.getTextFrame().getTextRange();

//Set the font
textRange.setLatinFont(new TextFont("Times New Roman"));

//Set bold property of the font
textRange.isBold(TriState.TRUE);

//Set italic property of the font
textRange.isItalic(TriState.TRUE);

//Set underline property of the font
textRange.setTextUnderlineType(TextUnderlineType.SINGLE);

//Set the height of the font
textRange.setFontHeight(50);

//Set the color of the font
textRange.getFill().setFillType(FillFormatType.SOLID);
textRange.getFill().getSolidColor().setColor(Color.cyan);
```

---

# Spire.Presentation Text Margins
## Set text margins for shapes in PowerPoint presentation
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Append a new shape
IAutoShape shape = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle(50, 100, 450, 150));

//Set the margins for the text frame
shape.getTextFrame().setMarginTop(10);
shape.getTextFrame().setMarginBottom(35);
shape.getTextFrame().setMarginLeft(15);
shape.getTextFrame().setMarginRight(30);
```

---

# Spire.Presentation Text Transparency
## Set text transparency with different alpha values in a presentation
```java
//Add a shape
IAutoShape textboxShape = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle(100, 100, 300, 120));
textboxShape.getShapeStyle().getLineColor().setColor(Color.white);
textboxShape.getFill().setFillType(FillFormatType.NONE);

//Remove default blank paragraphs
textboxShape.getTextFrame().getParagraphs().clear();

//Add three paragraphs, apply color with different alpha values to text
float alpha = 0.25f;
for (int i = 0; i < 3; i++) {
    textboxShape.getTextFrame().getParagraphs().append(new ParagraphEx());
    textboxShape.getTextFrame().getParagraphs().get(i).getTextRanges().append(new PortionEx("Text Transparency"));
    textboxShape.getTextFrame().getParagraphs().get(i).getTextRanges().get(0).getFill().setFillType(FillFormatType.SOLID);
    Color color = new Color(1.0F, 0.75F, 0.0F, alpha);
    textboxShape.getTextFrame().getParagraphs().get(i).getTextRanges().get(0).getFill().getSolidColor().
            setColor(color);
    alpha += 0.2;
}
```

---

# Spire.Presentation superscript and subscript
## Create superscript and subscript text in a presentation
```java
//Add a shape for superscript text
IAutoShape shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle(150, 100, 200, 50));
shape.getFill().setFillType(FillFormatType.NONE);
shape.getShapeStyle().getLineColor().setColor(Color.white);
shape.getTextFrame().getParagraphs().clear();

shape.appendTextFrame("Test");
PortionEx tr = new PortionEx("superscript");
shape.getTextFrame().getParagraphs().get(0).getTextRanges().append(tr);

//Set superscript text
shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(1).getFormat().setScriptDistance(30);

PortionEx textRange = shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0);
textRange.getFill().setFillType(FillFormatType.SOLID);
textRange.getFill().getSolidColor().setColor(Color.BLACK);
textRange.setFontHeight(20);
textRange.setLatinFont(new TextFont("Arial"));

textRange = shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(1);
textRange.getFill().setFillType(FillFormatType.SOLID);
textRange.getFill().getSolidColor().setColor(Color.BLUE);
textRange.setLatinFont(new TextFont("Arial"));

//Add a shape for subscript text
shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle(150, 150, 200, 50));
shape.getFill().setFillType(FillFormatType.NONE);
shape.getShapeStyle().getLineColor().setColor(Color.white);
shape.getTextFrame().getParagraphs().clear();

shape.appendTextFrame("Test");
tr = new PortionEx("subscript");
shape.getTextFrame().getParagraphs().get(0).getTextRanges().append(tr);

//Set subscript text
shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(1).getFormat().setScriptDistance(-25);

textRange = shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0);
textRange.getFill().setFillType(FillFormatType.SOLID);
textRange.getFill().getSolidColor().setColor(Color.BLACK);
textRange.setFontHeight(20);
textRange.setLatinFont(new TextFont("Arial"));

textRange = shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(1);
textRange.getFill().setFillType(FillFormatType.SOLID);
textRange.getFill().getSolidColor().setColor(Color.BLUE);
textRange.setLatinFont(new TextFont("Arial"));
```

---

# Spire Presentation Image in Master Slide
## Add an image to a master slide in a PowerPoint presentation

```java
//Get the master collection
IMasterSlide master = presentation.getMasters().get(0);

//Append image to slide master
Rectangle2D.Double rff = new Rectangle2D.Double(40, 40, 90, 90);
String imageFile = "data/logo.png";
IEmbedImage pic = master.getShapes().appendEmbedImage(ShapeType.RECTANGLE, imageFile, rff);
pic.getLine().getFillFormat().setFillType(FillFormatType.NONE);
//Add new slide to presentation
presentation.getSlides().append();
```

---

# Spire Presentation Slide Management
## Add slides using master layout
```java
//get Master layouts
ILayout iLayout = presentation.getMasters().get(0).getLayouts().get(0);

//append new slide
presentation.getSlides().append(iLayout);

//insert new slide
presentation.getSlides().insert(1, iLayout);
```

---

# Spire.Presentation slide management
## Append slides with master layouts
```java
//Get the master
IMasterSlide master = presentation.getMasters().get(0);
//Get master layout slides
IMasterLayouts masterLayouts = master.getLayouts();
ActiveSlide layoutSlide = (ActiveSlide) ((masterLayouts.get(1) instanceof ActiveSlide) ? masterLayouts.get(1) : null);

//Append a rectangle to the layout slide
IAutoShape shape = layoutSlide.getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle2D.Double(10, 50, 100, 80));
//Add a text into the shape and set the style
shape.getFill().setFillType(FillFormatType.NONE);
shape.appendTextFrame("Layout slide 1");
shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0).setLatinFont(new TextFont("Arial Black"));
shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0).getFill().setFillType(FillFormatType.SOLID);
shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0).getFill().getSolidColor().setColor(Color.orange);

//Append new slide with master layout
presentation.getSlides().append(presentation.getSlides().get(0), master.getLayouts().get(1));
//Another way to append new slide with master layout
presentation.getSlides().insert(2, presentation.getSlides().get(1), master.getLayouts().get(1));
```

---

# Spire.Presentation Slide Master
## Apply custom settings to a slide master in PowerPoint presentation
```java
//Get the first slide master from the presentation
IMasterSlide masterSlide = ppt.getMasters().get(0);

//Customize the background of the slide master
Rectangle rect = new Rectangle(0, 0, (int) ppt.getSlideSize().getSize().getWidth(), (int) ppt.getSlideSize().getSize().getHeight());
masterSlide.getSlideBackground().getFill().setFillType(FillFormatType.PICTURE);
masterSlide.getShapes().appendEmbedImage(ShapeType.RECTANGLE, "background_image_path", rect);
masterSlide.getSlideBackground().getFill().getPictureFill().getPicture().setUrl("background_image_path");

//Change the color scheme
masterSlide.getTheme().getColorScheme().getAccent1().setColor(Color.red);
masterSlide.getTheme().getColorScheme().getAccent2().setColor(Color.cyan);
masterSlide.getTheme().getColorScheme().getAccent3().setColor(Color.orange);
masterSlide.getTheme().getColorScheme().getAccent4().setColor(Color.BLACK);
```

---

# Spire.Presentation Slide Position Management
## Change slide position in presentation
```java
//Move the first slide to the second slide position
ISlide slide = ppt.getSlides().get(0);
slide.setSlideNumber(2);
```

---

# Spire Presentation Slide Cloning
## Clone slides from one presentation and append them to another presentation
```java
//Loop through all slides of source document
for (Object s : sourcePPT.getSlides()) {
    ISlide slide = (ISlide) s;
    //Append the slide at the end of destination document
    destPPT.getSlides().append(slide);
}
```

---

# Spire Presentation Master Slide Cloning
## Clone master slides from one PowerPoint presentation to another
```java
// Create presentation objects
Presentation presentation1 = new Presentation();
Presentation presentation2 = new Presentation();

// Clone masters from PPT1 to PPT2
for (Object obj : presentation1.getMasters()) {
    IMasterSlide masterSlide = (IMasterSlide) obj;
    presentation2.getMasters().appendSlide(masterSlide);
}
```

---

# Slide Cloning with Content Adaptation
## This code demonstrates how to clone slides between presentations with adaptive size adjustment.
```java
//Set the adaptive size when cloning slide, currently only supports 4:3->16:9
presentation1.isSlideSizeAutoFit(true);
ILayout layout = presentation1.getSlides().get(0).getLayout();
presentation1.getSlides().append(presentation2.getSlides().get(0), layout);
```

---

# Spire Presentation Slide Cloning
## Clone a slide and append it at the end of the presentation
```java
//Get the first slide
ISlide slide = presentation.getSlides().get(0);

//Append the slide at the end of the document
presentation.getSlides().append(slide);
```

---

# spire presentation slide cloning
## clone slide from one presentation to another
```java
// Create two presentations
Presentation presentation = new Presentation();
Presentation ppt1 = new Presentation();

// Get the first slide from the second presentation
ISlide slide1 = ppt1.getSlides().get(0);

// Insert the slide to the specified index in the source presentation
int index = 1;
presentation.getSlides().insert(index, slide1);
```

---

# Spire Presentation Slide Cloning
## Clone a slide within the same PowerPoint presentation
```java
//Get a list of slides and choose the first slide to be cloned
ISlide slide = ppt.getSlides().get(0);

//Insert the desired slide to the specified index in the same presentation
int index = 1;
ppt.getSlides().insert(index, slide);
```

---

# Spire Presentation Slide Creation
## Create and format slides with shapes and text
```java
// Create a new presentation
Presentation ppt = new Presentation();

// Add new slide
ppt.getSlides().append();

// Set the background image for slides
for (int i = 0; i < 2; i++) {
    Rectangle2D.Double rect = new Rectangle2D.Double(0, 0, ppt.getSlideSize().getSize().getWidth(), ppt.getSlideSize().getSize().getHeight());
    ppt.getSlides().get(i).getShapes().appendEmbedImage(ShapeType.RECTANGLE, imageFile, rect);
    ppt.getSlides().get(i).getShapes().get(0).getLine().getFillFormat().getSolidFillColor().setColor(Color.white);
}

// Add title shape
Rectangle2D.Double rec_title = new Rectangle2D.Double(ppt.getSlideSize().getSize().getWidth() / 2 - 200, 70, 400, 50);
IAutoShape shape_title = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, rec_title);
shape_title.getShapeStyle().getLineColor().setColor(Color.white);
shape_title.getFill().setFillType(FillFormatType.NONE);

// Add title paragraph
ParagraphEx para_title = new ParagraphEx();
para_title.setAlignment(TextAlignmentType.CENTER);
para_title.getTextRanges().get(0).setLatinFont(new TextFont("Myriad Pro Light"));
para_title.getTextRanges().get(0).setFontHeight(36);
para_title.getTextRanges().get(0).getFill().setFillType(FillFormatType.SOLID);
para_title.getTextRanges().get(0).getFill().getSolidColor().setColor(Color.darkGray);
shape_title.getTextFrame().getParagraphs().append(para_title);

// Append new shape
IAutoShape shape = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle2D.Double(50, 150, 600, 280));
shape.getShapeStyle().getLineColor().setColor(Color.white);
shape.getFill().setFillType(FillFormatType.NONE);
shape.getLine().setFillType(FillFormatType.NONE);

// Add text frame to shape
shape.appendTextFrame("Welcome to use Spire.Presentation for JAVA");

// Add new paragraph
ParagraphEx pare = new ParagraphEx();
pare.setText("");
shape.getTextFrame().getParagraphs().append(pare);

// Add another paragraph with description
pare = new ParagraphEx();
pare.setText("Spire.Presentation for Java is a professional PowerPoint API...");
shape.getTextFrame().getParagraphs().append(pare);

// Set font properties for paragraphs
for (Object para : shape.getTextFrame().getParagraphs()) {
    ((ParagraphEx) para).getTextRanges().get(0).setLatinFont(new TextFont("Myriad Pro"));
    ((ParagraphEx) para).getTextRanges().get(0).setFontHeight(24);
    ((ParagraphEx) para).getTextRanges().get(0).getFill().setFillType(FillFormatType.SOLID);
    ((ParagraphEx) para).getTextRanges().get(0).getFill().getSolidColor().setColor(Color.darkGray);
    ((ParagraphEx) para).setAlignment(TextAlignmentType.LEFT);
}

// Append star shape to second slide
shape = ppt.getSlides().get(1).getShapes().appendShape(ShapeType.SIX_POINTED_STAR, new Rectangle2D.Double(100, 100, 100, 100));
shape.getFill().setFillType(FillFormatType.SOLID);
shape.getFill().getSolidColor().setColor(Color.orange);
shape.getShapeStyle().getLineColor().setColor(Color.white);

// Append rectangle shape to second slide
shape = ppt.getSlides().get(1).getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle2D.Double(50, 250, 600, 50));
shape.getShapeStyle().getLineColor().setColor(Color.white);
shape.getFill().setFillType(FillFormatType.NONE);

// Add text to shape
shape.appendTextFrame("This is a newly added Slide.");

// Set font properties
shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0).setLatinFont(new TextFont("Myriad Pro"));
shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0).setFontHeight(24);
shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0).getFill().setFillType(FillFormatType.SOLID);
shape.getTextFrame().getParagraphs().get(0).getTextRanges().get(0).getFill().getSolidColor().setColor(Color.black);
shape.getTextFrame().getParagraphs().get(0).setAlignment(TextAlignmentType.LEFT);
shape.getTextFrame().getParagraphs().get(0).setIndent(35);
```

---

# Spire Presentation Slide Master
## Create and apply slide masters with background images
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();

ppt.getSlideSize().setType(SlideSizeType.SCREEN_16_X_9);

//Add slides
for (int i = 0; i < 4; i++) {
    ppt.getSlides().append();
}

//Get the first default slide master
IMasterSlide first_master = ppt.getMasters().get(0);

//Append another slide master
ppt.getMasters().appendSlide(first_master);
IMasterSlide second_master = ppt.getMasters().get(1);

//Set different background image for the two slide masters
//The first slide master
Rectangle rect = new Rectangle(0, 0, (int) ppt.getSlideSize().getSize().getWidth(), (int) ppt.getSlideSize().getSize().getHeight());
first_master.getSlideBackground().getFill().setFillType(FillFormatType.PICTURE);
first_master.getShapes().appendEmbedImage(ShapeType.RECTANGLE, "first_image_path", rect);
first_master.getSlideBackground().getFill().getPictureFill().getPicture().setUrl("first_image_path");
//The second slide master
second_master.getSlideBackground().getFill().setFillType(FillFormatType.PICTURE);
second_master.getShapes().appendEmbedImage(ShapeType.RECTANGLE, "second_image_path", rect);
second_master.getSlideBackground().getFill().getPictureFill().getPicture().setUrl("second_image_path");

//Apply the first master with layout to the first slide
ppt.getSlides().get(0).setLayout(first_master.getLayouts().get(1));

//Apply the second master with layout to other slides
for (int i = 1; i < ppt.getSlides().getCount(); i++) {
    ppt.getSlides().get(i).setLayout(second_master.getLayouts().get(8));
}
```

---

# Detect Used Themes in Presentation
## Extract theme names from each slide in a PowerPoint presentation
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();

StringBuilder sb = new StringBuilder();
String themeName = null;
sb.append("This is the name list of the used theme below.\r\t");
//Get the theme name of each slide in the document
for (Object obj : ppt.getSlides()) {
    ISlide slide = (ISlide) obj;
    themeName = slide.getTheme().getName();
    sb.append(themeName + "\r\t");
}
```

---

# Spire.Presentation Get Color Map
## Extract color mapping from PowerPoint presentation master slide
```java
//Create a PPT document
Presentation presentation = new Presentation();
IMasterSlide masterSlide = presentation.getMasters().get(0);
for (SchemeColor schemeColor : masterSlide.getColorMap().keySet()) {
    masterSlide.getColorMap().get(schemeColor);
}
```

---

# Spire Presentation Slide Access
## Get slides by index or ID and add shapes to them
```java
//Get slide by index 0
ISlide slide1 = presentation.getSlides().get(0);
//Append a shape in the slide
IAutoShape shape1=slide1.getShapes().appendShape(ShapeType.RECTANGLE, new java.awt.geom.Rectangle2D.Double(100, 100, 200, 100));
//Add text in the shape
shape1.getTextFrame().setText("Get slide by index");

//Get slide by slide ID
ISlide slide2 = presentation.findSlide((int)presentation.getSlides().get(1).getSlideID());
//Append a shape in the slide
IAutoShape shape2 = slide2.getShapes().appendShape(ShapeType.RECTANGLE, new java.awt.geom.Rectangle2D.Double(100, 100, 200, 100));
//Add text in the shape
shape2.getTextFrame().setText("Get slide by slide id");
```

---

# Spire Presentation Slide Layout Name
## Get the names of slide layouts in a PowerPoint presentation
```java
//Create a PPT document
Presentation presentation=new Presentation();

//Load a presentation from file
presentation.loadFromFile("presentation.pptx");

StringBuilder builder = new StringBuilder();

//Loop through the slides of PPT document
for (int i = 0; i < presentation.getSlides().getCount(); i++)
{
    //Get the name of slide layout
    String name = presentation.getSlides().get(i).getLayout().getName();
    builder.append(String.format("The name of slide %d layout is: %s", i,name)+"\r\n");
}
```

---

# Spire Presentation Text Extraction
## Extract text from presentation slides
```java
//Load PPT from disk
Presentation ppt = new Presentation();
ppt.loadFromFile("data/GetSlideText.pptx");
//Loop through all slides
for(int i=0;i<ppt.getSlides().getCount();i++){
    //Get text from slide
    ArrayList arrayList=ppt.getSlides().get(i).getAllTextFrame();
    for(int j=0;j<arrayList.size();j++){
        //Text extraction from slide
    }
}
```

---

# Spire Presentation Hide Slide
## Hide a specific slide in a PowerPoint presentation
```java
//Create a PPT document
Presentation ppt = new Presentation();

//Hide the second slide
ppt.getSlides().get(1).setHidden(true);
```

---

# Merge Selected Slides
## Merge slides from multiple presentations into a single presentation
```java
//Append all slides in ppt1 to ppt
for (int i = 0; i < ppt1.getSlides().getCount(); i++)
{
    ppt.getSlides().append(ppt1.getSlides().get(i));
}

//Append the second slide in ppt2 to ppt
ppt.getSlides().append(ppt2.getSlides().get(1));
```

---

# Spire.Presentation Slide Removal
## Remove slides from a presentation by index or reference
```java
//Remove slide by index
presentation.getSlides().removeAt(0);

//Remove the slide by its reference
ISlide slide = presentation.getSlides().get(1);
presentation.getSlides().remove(slide);
```

---

# Spire.Presentation Remove Unused Layout Master
## This code demonstrates how to remove unused layout masters from a PowerPoint presentation
```java
//Create an array list
ArrayList list = new ArrayList();
for (int i = 0; i < ppt.getSlides().getCount(); i++) {
    //Get the layout used by slide
    ActiveSlide layout = (ActiveSlide)ppt.getSlides().get(i).getLayout();
    list.add(layout);
}

//Loop through masters and layouts
for (int i = 0;i<ppt.getMasters().getCount(); i++) {
    IMasterLayouts masterlayouts = ppt.getMasters().get(i).getLayouts();
    for (int j=masterlayouts.getCount()-1;j>=0;j--)
    {
        if (!list.contains(masterlayouts.get(j)))
        {
            //Remove unused layout
            masterlayouts.removeMasterLayout(j);
        }
    }
}
```

---

# spire presentation slide numbering
## Set starting number for slides in PowerPoint presentation
```java
//Set 5 as the starting number
presentation.setFirstSlideNumber(5);
```

---

# Spire Presentation Slide Title Management
## Get and set slide titles in a PowerPoint presentation
```java
//Create PPT document
Presentation presentation = new Presentation();

//Get the first slide
ISlide slide = presentation.getSlides().get(0);

//Get the title of the first slide
String slideTitle = slide.getTitle();

//Set the title of the second slide
presentation.getSlides().get(1).setTitle("Second Slide");
```

---

# Java Presentation Slides Layout
## Add different layout slides to a presentation
```java
//Create a PPT document
Presentation presentation = new Presentation();

//Remove the default slide
presentation.getSlides().removeAt(0);

//Loop through slide layouts
for (SlideLayoutType type : SlideLayoutType.values())
{
    //Append slide by specifing slide layout
    presentation.getSlides().append(type);
}
```

---

# spire presentation slide layout
## change slide layout in presentation
```java
//Change the layout of slide
ppt.getSlides().get(1).setLayout(ppt.getMasters().get(0).getLayouts().get(4));
```

---

# spire presentation slide layout
## set slide layout in presentation
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Remove the first slide
ppt.getSlides().removeAt(0);

//Append a slide and set the layout for slide
ISlide slide = ppt.getSlides().append(SlideLayoutType.TITLE);
```

---

# Spire.Presentation Slide Transitions
## Set different transition types and timing for presentation slides
```java
//Create PPT document
Presentation presentation = new Presentation();

//Set the first slide transition as circle
presentation.getSlides().get(0).getSlideShowTransition().setType(TransitionType.CIRCLE);

//Set the transition time of 3 seconds
presentation.getSlides().get(0).getSlideShowTransition().setAdvanceOnClick(true);
presentation.getSlides().get(0).getSlideShowTransition().setAdvanceAfterTime(3000);

//Set the second slide transition as comb and set the speed
presentation.getSlides().get(1).getSlideShowTransition().setType(TransitionType.COMB);
presentation.getSlides().get(1).getSlideShowTransition().setSpeed(TransitionSpeed.SLOW);

// Set the transition time of 5 seconds
presentation.getSlides().get(1).getSlideShowTransition().setAdvanceOnClick(true);
presentation.getSlides().get(1).getSlideShowTransition().setAdvanceAfterTime(5000);

// Set the third slide transition as zoom
presentation.getSlides().get(2).getSlideShowTransition().setType(TransitionType.ZOOM);

// Set the transition time of 7 seconds
presentation.getSlides().get(2).getSlideShowTransition().setAdvanceOnClick(true);
presentation.getSlides().get(2).getSlideShowTransition().setAdvanceAfterTime(7000);
```

---

# Spire.Presentation Slide Transition
## Set advance time for slide transitions
```java
//Traverse all slides
for(int i=0;i<ppt.getSlides().getCount();i++)
{ 
    ppt.getSlides().get(i).getSlideShowTransition().isAdvanceAfterTime(true);

    //Set the time
    ppt.getSlides().get(i).getSlideShowTransition().setAdvanceAfterTime(5000L);
}
```

---

# Spire Presentation Transition Effects
## Set slide transition effects in a presentation
```java
// Create PPT document
Presentation presentation = new Presentation();

// Set effects
presentation.getSlides().get(0).getSlideShowTransition().setType(TransitionType.CUT);
((OptionalBlackTransition)presentation.getSlides().get(0).getSlideShowTransition().getValue()).setFromBlack(true);
```

---

# Spire Presentation Slide Transitions
## Set slide transitions in a presentation
```java
//Set the first slide transition as push and sound mode
presentation.getSlides().get(0).getSlideShowTransition().setType(TransitionType.PUSH);
presentation.getSlides().get(0).getSlideShowTransition().setSoundMode(TransitionSoundMode.START_SOUND);

//Set the second slide transition as circle and set the speed
presentation.getSlides().get(1).getSlideShowTransition().setType(TransitionType.FADE);
presentation.getSlides().get(1).getSlideShowTransition().setSpeed(TransitionSpeed.SLOW);
```

---

# Add Line to PowerPoint Slide
## This code demonstrates how to add a line shape to a PowerPoint slide and customize its appearance
```java
//Create a PPT document
Presentation presentation = new Presentation();

//Get the first slide
ISlide slide = presentation.getSlides().get(0);

//Add a line in the slide
IAutoShape line=slide.getShapes().appendShape(ShapeType.LINE, new Rectangle2D.Double(50, 100, 300, 0));

//Set color of the line
line.getShapeStyle().getLineColor().setColor(Color.red);
```

---

# Add Lines with Arrows in Presentation
## Create lines with different arrow styles and colors in a presentation slide
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Add a line to the slides and set its color to red
IAutoShape shape = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.LINE, new Rectangle2D.Float(150, 100, 100, 100));
shape.getShapeStyle().getLineColor().setColor(Color.red);
//Set the line end type as StealthArrow
shape.getLine().setLineEndType(LineEndType.STEALTH_ARROW);

//Add a line to the slides and use default color
shape = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.LINE, new Rectangle2D.Float(300, 150, 100, 100));
shape.setRotation(-45);
//Set the line end type as TriangleArrowHead
shape.getLine().setLineEndType(LineEndType.TRIANGLE_ARROW_HEAD);

//Add a line to the slides and set its color to green
shape = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.LINE, new Rectangle2D.Float(450, 100, 100, 100));
shape.getShapeStyle().getLineColor().setColor(Color.green);
shape.setRotation(90);
//Set the line begin type as TriangleArrowHead
shape.getLine().setLineEndType(LineEndType.STEALTH_ARROW);
```

---

# Spire Presentation MathML Equation
## Add MathML equation to PowerPoint slide
```java
//Create a PPT document
Presentation ppt = new Presentation();

//Set the mathML code
String mathMLCode="<mml:math xmlns:mml=\"http://www.w3.org/1998/Math/MathML\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">" + "<mml:msup><mml:mrow><mml:mi>x</mml:mi></mml:mrow><mml:mrow><mml:mn>2</mml:mn></mml:mrow></mml:msup><mml:mo>+</mml:mo><mml:msqrt><mml:msup><mml:mrow><mml:mi>x</mml:mi></mml:mrow><mml:mrow><mml:mn>2</mml:mn></mml:mrow></mml:msup><mml:mo>+</mml:mo><mml:mn>1</mml:mn></mml:msqrt><mml:mo>+</mml:mo><mml:mn>1</mml:mn></mml:math>";

//Add a shape
IAutoShape shape=ppt.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle2D.Float(30,100,400,30));
shape.getTextFrame().getParagraphs().clear();

//Add the mathml equation paragraph
ParagraphEx tp = shape.getTextFrame().getParagraphs().addParagraphFromMathMLCode(mathMLCode);
```

---

# Spire.Presentation Round Corner Rectangle
## Add a round corner rectangle shape to a presentation slide
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Append a round corner rectangle and set its radius
IAutoShape shape = ppt.getSlides().get(0).getShapes().appendRoundRectangle(300, 90, 100, 200, 80);
//Set the color and fill style of shape
shape.getFill().setFillType(FillFormatType.SOLID);
shape.getFill().getSolidColor().setColor(Color.PINK);
shape.getShapeStyle().getLineColor().setColor(Color.LIGHT_GRAY);
//Rotate the shape to 90 degree
shape.setRotation(90);
```

---

# Spire.Presentation Shape Addition
## Add various shapes to a PowerPoint presentation
```java
//Create PPT document
Presentation presentation = new Presentation();

//Append new shape - Triangle and set style
IAutoShape shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.TRIANGLE, new Rectangle2D.Double(115, 130, 100, 100));

shape.getFill().setFillType(FillFormatType.SOLID);
shape.getFill().getSolidColor().setColor(Color.green);
shape.getShapeStyle().getLineColor().setColor(Color.white);

//Append new shape - Ellipse
shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.ELLIPSE, new Rectangle2D.Double(290, 130, 150, 100));
shape.getFill().setFillType(FillFormatType.SOLID);
shape.getFill().getSolidColor().setColor(Color.blue);
shape.getShapeStyle().getLineColor().setColor(Color.white);

//Append new shape - Heart
shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.HEART, new Rectangle2D.Double(470, 130, 130, 100));
shape.getFill().setFillType(FillFormatType.SOLID);
shape.getFill().getSolidColor().setColor(Color.red);
shape.getShapeStyle().getLineColor().setColor(Color.gray);

//Append new shape - FivePointedStar
shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.FIVE_POINTED_STAR, new Rectangle2D.Double(90, 270, 150, 150));
shape.getFill().setFillType(FillFormatType.GRADIENT);
shape.getFill().getSolidColor().setColor(Color.black);
shape.getShapeStyle().getLineColor().setColor(Color.white);

//Append new shape - Rectangle
shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.TRIANGLE, new Rectangle2D.Double(320, 290, 100, 120));
shape.getFill().setFillType(FillFormatType.SOLID);
shape.getFill().getSolidColor().setColor(Color.pink);
shape.getShapeStyle().getLineColor().setColor(Color.lightGray);

//Append new shape - BentUpArrow
shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.BENT_UP_ARROW, new Rectangle2D.Double(470, 300, 150, 100));
//Set the color of the shape
shape.getFill().setFillType(FillFormatType.GRADIENT);
shape.getFill().getGradient().getGradientStops().append(1f, KnownColors.OLIVE);
shape.getFill().getGradient().getGradientStops().append(0, KnownColors.POWDER_BLUE);
shape.getShapeStyle().getLineColor().setColor(Color.white);
```

---

# spire presentation image stream
## append image from stream to presentation slide
```java
// Create presentation object
Presentation ppt = new Presentation();

// Load image from file stream
FileInputStream fileInputStream = new FileInputStream(intputFile_Img);
IImageData imageData = ppt.getImages().append(fileInputStream);

// Get the first shape on the first slide and replace its image
SlidePicture slidePicture = (SlidePicture) ppt.getSlides().get(0).getShapes().get(0);
slidePicture.getPictureFill().getPicture().setEmbedImage(imageData);
```

---

# Shape Arrangement in Presentation
## Demonstrates how to arrange shapes in a PowerPoint presentation by bringing a shape forward
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Get the specified shape
IShape shape = ppt.getSlides().get(0).getShapes().get(0);

//Bring the shape forward through SetShapeArrange method
shape.setShapeArrange(ShapeAlignmentEnum.ShapeArrange.BringForward);
```

---

# spire.presentation background
## Set background image for presentation slide
```java
//Set background Image
Rectangle2D.Double rect = new Rectangle2D.Double(0, 0, presentation.getSlideSize().getSize().getWidth(), presentation.getSlideSize().getSize().getHeight());
presentation.getSlides().get(0).getShapes().appendEmbedImage(ShapeType.RECTANGLE, imageFile, rect);
```

---

# Convert Group Shape to Image
## Convert group shapes in presentation slides to image files
```java
Presentation ppt = new Presentation();
ppt.loadFromFile("input.pptx");

// Convert group shape to image
for (int i = 0; i < ppt.getSlides().get(0).getShapes().getCount(); i++){
    String fileName = "shapeToImage_" + i + ".png";
    //Convert shape to image
    BufferedImage image = ppt.getSlides().get(0).getShapes().saveAsImage(i);
    ImageIO.write(image, "PNG", new File(fileName));
}
```

---

# Spire.Presentation Shape Copying
## Copy shapes between slides in a presentation
```java
//Define the source slide and target slide
ISlide sourceSlide = ppt.getSlides().get(0);
ISlide targetSlide = ppt.getSlides().get(1);

//Copy the first shape from the source slide to the target slide
targetSlide.getShapes().addShape((Shape)sourceSlide.getShapes().get(0));
```

---

# Creating Irregular Polygon in PowerPoint
## This code demonstrates how to create an irregular polygon shape in a PowerPoint presentation.
```java
// Create a new PowerPoint presentation
Presentation ppt = new Presentation();

// Get the first slide of the presentation
ISlide slide = ppt.getSlides().get(0);

// Define the points for the irregular polygon shape
List<Point2D> points = new ArrayList<>();
points.add(new Point2D.Float(50f, 50f));
points.add(new Point2D.Float(50f, 150f));
points.add(new Point2D.Float(60f, 200f));
points.add(new Point2D.Float(200f, 200f));
points.add(new Point2D.Float(220f, 150f));
points.add(new Point2D.Float(150f, 90f));
points.add(new Point2D.Float(50f, 50f));

// Append the irregular polygon shape to the slide
IAutoShape autoShape = slide.getShapes().appendFreeformShape(points);

// Set the fill type of the shape to none (transparent)
autoShape.getFill().setFillType(FillFormatType.NONE);
```

---

# Spire.Presentation Line Drawing
## Draw a line between two points on a PowerPoint slide
```java
// Create a new PowerPoint presentation
Presentation ppt = new Presentation();

// Get the first slide of the presentation
ISlide slide = ppt.getSlides().get(0);

// Define the starting point and ending point for the line
Point2D startPoint = new Point2D.Float(50, 70);
Point2D endPoint = new Point2D.Float(150, 120);

// Append a line shape between the specified points to the slide
slide.getShapes().appendShape(ShapeType.LINE, startPoint, endPoint);
```

---

# Presentation Animation Duration and Delay Time Control
## Control animation duration and delay time in a presentation
```java
//Get the first slide
ISlide slide = presentation.getSlides().get(0);
AnimationEffectCollection animations = slide.getTimeline().getMainSequence();

//Get duration time of animation
float durationTime = animations.get(0).getTiming().getDuration();

//Set new duration time of animation
animations.get(0).getTiming().setDuration(0.8f);

//Get delay time of animation
float delayTime = animations.get(0).getTiming().getTriggerDelayTime();

//Set new delay time of animation
animations.get(0).getTiming().setTriggerDelayTime(0.6f);
```

---

# Java Presentation Gradient Shape Fill
## Fill a shape with gradient colors in a PowerPoint presentation
```java
//Get the first shape and set the style to be Gradient
IAutoShape GradientShape = (IAutoShape) ppt.getSlides().get(0).getShapes().get(0);
GradientShape.getFill().setFillType(FillFormatType.GRADIENT);
GradientShape.getFill().getGradient().getGradientStops().append(0, Color.blue);
GradientShape.getFill().getGradient().getGradientStops().append(1, Color.lightGray);
```

---

# Spire Presentation Pattern Fill
## Fill a shape with pattern in PowerPoint presentation
```java
//Create a PPT document
Presentation presentation = new Presentation();

//Get the first slide
ISlide slide = presentation.getSlides().get(0);

//Add a rectangle
Rectangle2D rect = new Rectangle2D.Double(presentation.getSlideSize().getSize().getWidth() / 2 - 50, 100, 100, 100);
IAutoShape shape = slide.getShapes().appendShape(ShapeType.RECTANGLE, rect);

//Set the pattern fill format
shape.getFill().setFillType( FillFormatType.PATTERN);
shape.getFill().getPattern().setPatternType( PatternFillType.TRELLIS);
shape.getFill().getPattern().getBackgroundColor().setColor( Color.darkGray);
shape.getFill().getPattern().getForegroundColor().setColor( Color.yellow);

//Set the fill format of line
shape.getLine().setFillType(FillFormatType.SOLID);
shape.getLine().getSolidFillColor().setColor(Color.white);
```

---

# Spire.Presentation Fill Shape With Picture
## Demonstrates how to fill a shape with a picture in a PowerPoint presentation
```java
//Get the first shape and set the style to be Gradient
IAutoShape shape = (IAutoShape) ppt.getSlides().get(0).getShapes().get(0);

//Fill the shape with picture
shape.getFill().setFillType(FillFormatType.PICTURE);
BufferedImage image = ImageIO.read( new File(imageURL));
shape.getFill().getPictureFill().getPicture().setEmbedImage(ppt.getImages().append(image));

shape.getFill().getPictureFill().setFillType(PictureFillType.STRETCH);
```

---

# Spire Presentation Shape Filling
## Fill shape with solid color in presentation
```java
//Create a PPT document
Presentation presentation = new Presentation();

//Get the first slide
ISlide slide = presentation.getSlides().get(0);

//Add a rectangle
Rectangle2D rect = new Rectangle2D.Double(presentation.getSlideSize().getSize().getWidth()/ 2 - 50, 100, 100, 100);
IAutoShape shape = slide.getShapes().appendShape(ShapeType.RECTANGLE, rect);

//Fill shape with solid color
shape.getFill().setFillType(FillFormatType.SOLID);
shape.getFill().getSolidColor().setColor(Color.yellow);

//Set the fill format of line
shape.getLine().setFillType(FillFormatType.SOLID);
shape.getLine().getSolidFillColor().setColor(Color.GRAY);
```

---

# Find Shape by Alternative Text
## Method to locate a shape in a presentation slide by its alternative text property
```java
private static IShape FindShape(ISlide slide, String altText)
{
    //Loop through shapes in the slide
    for (IShape shape : (Iterable<IShape>)slide.getShapes())
    {
        //Find the shape whose alternative text is altText
        if (shape.getAlternativeText().compareTo(altText) == 0)
        {
            return shape;
        }
    }
    return null;
}
```

---

# Spire.Presentation Get Titles
## Extract all title text from a presentation
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Instantiate a list of IShape objects
ArrayList<IShape> shapelist = new ArrayList<IShape>();

//Loop through all slides and all shapes on each slide
for (ISlide slide : (Iterable<ISlide>)ppt.getSlides())
{
    for (IShape shape :(Iterable<IShape>)slide.getShapes())
    {
        if (shape.getPlaceholder() != null)
        {
            //Get all titles
            switch (shape.getPlaceholder().getType())
            {
                case TITLE:
                    shapelist.add(shape);
                    break;
                case CENTERED_TITLE:
                    shapelist.add(shape);
                    break;
                case SUBTITLE:
                    shapelist.add(shape);
                    break;
            }
        }
    }
}

//Loop through the list and get the inner text of all shapes in the list
for (int i = 0; i < shapelist.size(); i++)
{
    IAutoShape shape1 = (IAutoShape)shapelist.get(i);
    shape1.getTextFrame().getText();
}
```

---

# Get Display Color of Shape
## This code demonstrates how to get the display color and fill type of a shape in a presentation
```java
// Get the first shape
IAutoShape shape = (IAutoShape)ppt.getSlides().get(0).getShapes().get(0);
// Get the fill type and color of the shape
shape.getDisplayFill().getFillType();
shape.getDisplayFill().getSolidColor().getColor();
```

---

# Spire.Presentation Layout Prototype
## Get layout prototype from a shape in a presentation
```java
//Get a shape from the first slide
IShape shape = presentation.getSlides().get(0).getShapes().get(0);

//Get layout prototype from the shape
Shape layoutPrototype = shape.getLayoutPrototype();
```

---

# Get Max Width of Text Area in PowerPoint Shape
## Demonstrates how to get the maximum width of a text area within a shape in a PowerPoint presentation
```java
// Get a slide from a presentation
ISlide slide = ppt.getSlides().get(0);
// Get a shape from the slide
IAutoShape shape = (IAutoShape)slide.getShapes().get(0);
// Get the max width of the text area of the shape
float maxWidth = shape.getTextFrame().getMaxWidth();
```

---

# Get Text Position in Presentation
## Extract the position of text within a slide and shape in a PowerPoint presentation
```java
ISlide slide = ppt.getSlides().get(0);
IAutoShape shape = (IAutoShape)slide.getShapes().get(0);
Point2D location = shape.getTextFrame().getTextLocation();
// Calculate text position relative to slide
double slideX = location.getX();
double slideY = location.getY();
// Calculate text position relative to shape
double shapeX = location.getX() - shape.getLeft();
double shapeY = location.getY() - shape.getTop();
```

---

# Spire.Presentation Shape Group Alternative Text
## Extract alternative text from shape groups in PowerPoint presentation
```java
//Create a PPT document
Presentation presentation = new Presentation();

StringBuilder builder=new StringBuilder();

//Loop through slides and shapes
for (ISlide slide : (Iterable<ISlide>)presentation.getSlides())
{
    for (IShape shape : (Iterable<IShape>)slide.getShapes())
    {
        if (shape instanceof GroupShape)
        {
            //Find the shape group
            GroupShape groupShape = (GroupShape)shape;
            int i=1;
            for (IShape gShape : (Iterable<IShape>)groupShape.getShapes())
            {
                //Append the alternative text in builder
                builder.append(gShape.getAlternativeText()+"\r\n");
            }
        }
    }
}
```

---

# Spire.Presentation Shape Points Extraction
## Get points from a shape in a PowerPoint presentation
```java
//Get the first shape in first slide
IAutoShape shape = (IAutoShape)ppt.getSlides().get(0).getShapes().get(0);

//Get the Point of shape
ArrayList<Point2D> points = shape.getPoints();
```

---

# Spire.Presentation Get Text Position
## Get the start position at maximum width of text area in a presentation shape
```java
ISlide slide = presentation.getSlides().get(0);
IAutoShape shape = (IAutoShape)slide.getShapes().get(0);
double x = shape.getTextFrame().getStartLocationAtMaxWidth().getX();
double y = shape.getTextFrame().getStartLocationAtMaxWidth().getY();
```

---

# Spire.Presentation Text Position in Shape
## Get text position information within AutoShapes in a PowerPoint slide
```java
// Access the first slide in the presentation
ISlide slide = ppt.getSlides().get(0);

// Iterate through all the shapes in the slide
for (int i = 0; i < slide.getShapes().getCount(); i++) {
    // Get the current shape
    IShape shape = slide.getShapes().get(i);

    // Check if the shape is an AutoShape
    if (shape instanceof IAutoShape) {
        // Cast the shape to an AutoShape
        IAutoShape autoshape = (IAutoShape) shape;

        // Get the text content of the AutoShape
        String text = autoshape.getTextFrame().getText();

        // Obtain the text position information within the AutoShape
        Point2D point = autoshape.getTextFrame().getTextLocation();
    }
}
```

---

# Spire.Presentation Shape Grouping
## Group shapes in a PowerPoint presentation
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Get the first slide
ISlide slide = ppt.getSlides().get(0);

//Create two shapes in the slide
IShape rectangle = slide.getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle2D.Double(250, 180, 200, 40));
rectangle.getFill().setFillType(FillFormatType.SOLID);
rectangle.getFill().getSolidColor().setKnownColor(KnownColors.LIGHT_BLUE);
rectangle.getLine().setWidth(0.1f);
IShape ribbon = slide.getShapes().appendShape(ShapeType.RIBBON_2, new Rectangle2D.Double(290, 155, 120, 80));
ribbon.getFill().setFillType(FillFormatType.SOLID);
ribbon.getFill().getSolidColor().setKnownColor(KnownColors.LIGHT_PINK);
ribbon.getLine().setWidth(0.1f);

//Add the two shape objects to an array list
ArrayList list = new ArrayList();
list.add(rectangle);
list.add(ribbon);

//Group the shapes in the list
ppt.getSlides().get(0).groupShapes(list);
```

---

# Spire.Presentation Hide Shape
## Hide a specific shape in a PowerPoint presentation by its alternative text
```java
//Loop through slides
for (ISlide slide : (Iterable<ISlide>) presentation.getSlides())
{
    //Loop through shapes in the slide
    for (IShape shape :(Iterable<IShape>) slide.getShapes())
    {
        //Find the shape whose alternative text is Shape1
        if (shape.getAlternativeText().compareTo("Shape1") == 0)
        {
            //Hide the shape
            shape.isHidden(true);
        }
    }
}
```

---

# PowerPoint Shape Textbox Detection
## Check if PowerPoint shapes are textboxes
```java
for (ISlide slide:(Iterable<? extends ISlide>) presentation.getSlides())
{
    for (IShape shape:(Iterable<? extends IShape>) slide.getShapes())
    {
        if (shape instanceof IAutoShape)
        {
            // Determine if the shape is a textbox
            boolean isTextbox = shape.isTextBox();
        }
    }
}
```

---

# Spire.Presentation Placeholder Operations
## Operate on different types of placeholders in a presentation
```java
//Operate placeholders
for (int j=0;j<presentation.getSlides().getCount();j++)
{
    ISlide slide = presentation.getSlides().get(j);

    for (int i=0;i<slide.getShapes().getCount();i++)
    {
        Shape shape = (Shape)slide.getShapes().get(i);
        switch(shape.getPlaceholder().getType())
        {
            case MEDIA:
                shape.insertVideo("data/Video.mp4");
                break;

            case PICTURE:
                shape.insertPicture("data/E-iceblueLogo.png");
                break;

            case CHART:
                shape.insertChart(ChartType.COLUMN_CLUSTERED);
                break;

            case TABLE:
                shape.insertTable(3,2);
                break;

            case DIAGRAM:
                shape.insertSmartArt(SmartArtLayoutType.BASIC_BLOCK_LIST);
                break;
        }
    }
}
```

---

# Spire.Presentation Shape Locking
## Prevent or allow changing shape properties in a presentation
```java
//The changes of selection and rotation are allowed
shape.getLocking().setRotationProtection(false);
shape.getLocking().setSelectionProtection(false);
//The changes of size, position, shape type, aspect ratio, text editing and ajust handles are not allowed
shape.getLocking().setResizeProtection(true);
shape.getLocking().setPositionProtection(true);
shape.getLocking().setShapeTypeProtection(true);
shape.getLocking().setAspectRatioProtection(true);
shape.getLocking().setTextEditingProtection(true);
shape.getLocking().setAdjustHandlesProtection(true);
```

---

# spire presentation remove shape
## remove shapes from presentation based on alternative text
```java
//Loop through slides
for (int i = 0; i < presentation.getSlides().getCount(); i++)
{
    ISlide slide = presentation.getSlides().get(i);
    //Loop through shapes
    for (int j = 0; j < slide.getShapes().getCount(); j++)
    {
        IShape shape = slide.getShapes().get(j);
        //Find the shapes whose alternative text contain "Shape"
        if(shape.getAlternativeText().contains("Shape"))
        {
            slide.getShapes().remove(shape);
            j--;
        }
    }
}
```

---

# Reorder Overlapping Shapes in Presentation
## Change the z-order of shapes to control overlapping
```java
//Get the first shape of the first slide
IShape shape = ppt.getSlides().get(0).getShapes().get(0);

//Change the shape's zorder
ppt.getSlides().get(0).getShapes().zOrder(1,shape);
```

---

# Reset Position of Placeholder
## Reset the position of slide number and date placeholders in a PowerPoint presentation
```java
//Get the first slide from the sample document.
ISlide slide = presentation.getSlides().get(0);

for (IShape shapeToMove :(Iterable<IShape>) slide.getShapes())
{
    //Reset the position of the slide number to the left.
    if (shapeToMove.getName().contains("Slide Number Placeholder"))
    {
        shapeToMove.setLeft(0);
    }

    else if (shapeToMove.getName().contains("Date Placeholder"))
    {
        //Reset the position of the date time to the center.
        shapeToMove.setLeft(presentation.getSlideSize().getSize().getWidth()/ 2);

        SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
        Date dt=new Date();
        String time=sf.format(dt);

        //Reset the date time display style.
        ((IAutoShape)shapeToMove).getTextFrame().getTextRange().getParagraph().setText(time);
        ((IAutoShape)shapeToMove).getTextFrame().isCentered(true);
    }
}
```

---

# Reset Shape Size and Position in Presentation
## Adjust shapes proportionally when changing slide size
```java
//Define the original slide size
double currentHeight = ppt.getSlideSize().getSize().getHeight();
double currentWidth = ppt.getSlideSize().getSize().getWidth();

//Change the slide size as A3
ppt.getSlideSize().setType(SlideSizeType.A3);

//Define the new slide size
double newHeight = ppt.getSlideSize().getSize().getHeight();
double newWidth = ppt.getSlideSize().getSize().getWidth();

//Define the ratio from the old and new slide size
double ratioHeight = newHeight / currentHeight;
double ratioWidth = newWidth / currentWidth;

//Reset the size and position of the shape on the slide
for (ISlide slide :(Iterable<ISlide>) ppt.getSlides())
{
    for (IShape shape: (Iterable<IShape>)slide.getShapes())
    {
        shape.setHeight((float) (shape.getHeight() * ratioHeight));
        shape.setWidth((float)(shape.getWidth() * ratioWidth));

        shape.setLeft((float)shape.getLeft() * ratioHeight);
        shape.setTop((float)shape.getTop() * ratioWidth);
    }
}
```

---

# Spire Presentation Shape Rotation
## Rotate shapes in a presentation by different angles
```java
//Get the shapes
IAutoShape shape = (IAutoShape) ppt.getSlides().get(0).getShapes().get(0);

//Set the rotation
shape.setRotation(60);

((IAutoShape) ppt.getSlides().get(0).getShapes().get(1)).setRotation(120);
((IAutoShape) ppt.getSlides().get(0).getShapes().get(2)).setRotation(180);
((IAutoShape) ppt.getSlides().get(0).getShapes().get(3)).setRotation(240);
```

---

# Spire.Presentation 3D Shape Effects
## Apply 3D effects to shapes in a presentation
```java
//Add shape1 and fill it with color
IAutoShape shape1 = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.ROUND_CORNER_RECTANGLE, new Rectangle2D.Double(150, 150, 150, 150));
shape1.getFill().setFillType(FillFormatType.SOLID);
shape1.getFill().getSolidColor().setKnownColor(KnownColors.SKY_BLUE);
//Initialize a new instance of the 3-D class for shape1 and set its properties
ShapeThreeD effect1 = shape1.getThreeD().getShapeThreeD();
effect1.setPresetMaterial(PresetMaterialType.POWDER);
effect1.getTopBevel().setPresetType(BevelPresetType.ART_DECO);
effect1.getTopBevel().setHeight(4);
effect1.getTopBevel().setWidth(12);
effect1.setBevelColorMode(BevelColorType.CONTOUR);
effect1.getContourColor().setKnownColor(KnownColors.LIGHT_GRAY);
effect1.setContourWidth(3.5);

//Add shape2 and fill it with color
IAutoShape shape2 = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.PENTAGON, new Rectangle2D.Double(400, 150, 150, 150));
shape2.getFill().setFillType(FillFormatType.SOLID);
shape2.getFill().getSolidColor().setKnownColor(KnownColors.LIGHT_GREEN);
//Initialize a new instance of the 3-D class for shape2 and set its properties
ShapeThreeD effect2 = shape2.getThreeD().getShapeThreeD();
effect2.setPresetMaterial(PresetMaterialType.SOFT_EDGE);
effect2.getTopBevel().setPresetType(BevelPresetType.SOFT_ROUND);
effect2.getTopBevel().setHeight(12);
effect2.getTopBevel().setWidth(12);
effect2.setBevelColorMode(BevelColorType.CONTOUR);
effect2.getContourColor().setKnownColor(KnownColors.LIGHT_GREEN);
effect2.setContourWidth(5);
```

---

# Spire.Presentation Alternative Text
## Set and get alternative text for shapes in a presentation
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Get the first slide
ISlide slide = ppt.getSlides().get(0);

//Set the alternative text (title and description)
slide.getShapes().get(0).setAlternativeTitle("Rectangle");
slide.getShapes().get(0).setAlternativeText("This is a Rectangle");

//Get the alternative text (title and description)
String alternativeText = null;
String title = slide.getShapes().get(0).getAlternativeTitle();
alternativeText += "Title: " + title + "\r\n";
String description = slide.getShapes().get(0).getAlternativeText();
alternativeText += "Description: " + description;
```

---

# PowerPoint Gradient Stop Transparency and Brightness
## Set brightness and transparency values for gradient stops in PowerPoint shapes
```java
// Get the first slide
ISlide slide = ppt.getSlides().get(0);

// Iterate through Shapes within a Slide
for (int j = 0; j < slide.getShapes().size(); j++) {
    // Get the specific shape
    IAutoShape shape = (IAutoShape) ((GroupShape) slide.getShapes().get(j)).getShapes().get(2);
    // Get the collection of gradient stops
    GradientStopCollection stops = shape.getFill().getGradient().getGradientStops();
    
    // Iterate through the collection of gradient stops
    for (int i = 0; i < stops.size(); i++) {
        // Get transparency and brightness
        float transparency = stops.get(i).getColor().getTransparency();
        float brightness = stops.get(i).getColor().getBrightness();
    }
    // Set transparency and brightness
    stops.get(0).getColor().setTransparency(0.1f);
    stops.get(0).getColor().setBrightness(-0.1f);
    stops.get(1).getColor().setTransparency(0.51f);
    stops.get(1).getColor().setBrightness(0.5f);
}
```

---

# Spire Presentation Ellipse Formatting
## Set fill and line formatting for an ellipse shape in a presentation
```java
//Create a PPT document
Presentation presentation = new Presentation();

//Get the first slide
ISlide slide = presentation.getSlides().get(0);

//Add a rectangle
Rectangle2D rect = new Rectangle2D.Double(presentation.getSlideSize().getSize().getWidth() / 2 - 100, 100, 200, 100);
IAutoShape shape = slide.getShapes().appendShape(ShapeType.ELLIPSE, rect);

//Set the fill format of shape
shape.getFill().setFillType(FillFormatType.SOLID);
shape.getFill().getSolidColor().setColor(Color.PINK);

//Set the fill format of line
shape.getLine().setFillType(FillFormatType.SOLID);
shape.getLine().getSolidFillColor().setColor(Color.gray);
```

---

# Java Presentation Line Formatting
## Set format for lines in presentation shapes
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Add a rectangle shape to the slide
IAutoShape shape = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle2D.Double(100, 150, 200, 100));
//Apply some formatting on the line of the rectangle
shape.getLine().setStyle(TextLineStyle.THICK_THIN);
shape.getLine().setWidth(5);
shape.getLine().setDashStyle(LineDashStyleType.DASH);
//Set the color of the line of the rectangle
shape.getShapeStyle().getLineColor().setColor(Color.blue);

//Add a ellipse shape to the slide
shape = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.ELLIPSE, new Rectangle2D.Double(400, 150, 200, 100));
//Apply some formatting on the line of the ellipse
shape.getLine().setStyle(TextLineStyle.THICK_BETWEEN_THIN);
shape.getLine().setWidth(5);
shape.getLine().setDashStyle(LineDashStyleType.DASH_DOT);
//Set the color of the line of the ellipse
shape.getShapeStyle().getLineColor().setColor(Color.orange);
```

---

# Spire.Presentation Line Join Styles
## Set different line join styles for shapes
```java
//Fill lines of shapes
shape1.getLine().setFillType(FillFormatType.SOLID);
shape1.getLine().getSolidFillColor().setColor(Color.gray);
shape2.getLine().setFillType(FillFormatType.SOLID);
shape2.getLine().getSolidFillColor().setColor(Color.gray);
shape3.getLine().setFillType(FillFormatType.SOLID);
shape3.getLine().getSolidFillColor().setColor(Color.gray);

//Set the line width
shape1.getLine().setWidth(10);
shape2.getLine().setWidth(10);
shape3.getLine().setWidth(10);

//Set the join styles of lines
shape1.getLine().setJoinStyle(LineJoinType.BEVEL);
shape2.getLine().setJoinStyle(LineJoinType.MITER);
shape3.getLine().setJoinStyle(LineJoinType.ROUND);
```

---

# Spire.Presentation Shape Outlines and Effects
## Demonstrates how to set outline colors and effects (shadow, glow) for shapes in a PowerPoint presentation
```java
//create an instance of presentation document
Presentation ppt = new Presentation();

//get the first slide
ISlide slide = ppt.getSlides().get(0);

//draw a Rectangle shape
IAutoShape shape = slide.getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle2D.Float(150, 180, 100, 50));
shape.getFill().setFillType(FillFormatType.SOLID);
shape.getFill().getSolidColor().setColor(Color.BLUE);

//set outline color
shape.getShapeStyle().getLineColor().setColor(Color.red);

//set shadow effect
PresetShadow shadow = new PresetShadow();
shadow.getColorFormat().setColor(Color.MAGENTA);
shadow.setPreset( PresetShadowValue.FRONT_RIGHT_PERSPECTIVE);
shadow.setDistance(10.0);
shadow.setDirection(225.0f);
shape.getEffectDag().setPresetShadowEffect(shadow);

//draw a Ellipse shape
shape = slide.getShapes().appendShape(ShapeType.ELLIPSE, new Rectangle2D.Float(400, 150, 100, 100));
shape.getFill().setFillType(FillFormatType.SOLID);
shape.getFill().getSolidColor().setColor(Color.BLUE);

//set outline color
shape.getShapeStyle().getLineColor().setColor(Color.YELLOW);

//set shadow effect
GlowEffect glow = new GlowEffect();
glow.getColorFormat().setColor(Color.PINK);
glow.setRadius(20.0);
shape.getEffectDag().setGlowEffect(glow);
```

---

# Spire.Presentation Rounded Rectangle Radius
## Set radius for different types of rounded rectangles in a presentation
```java
//Create a PPT document
Presentation presentation = new Presentation();
ISlide iSlide = presentation.getSlides().get(0);

//Insert a rectangle with four round corners and set its radius
IAutoShape autoShape1=iSlide.getShapes().appendShape(ShapeType.ROUND_CORNER_RECTANGLE,new Rectangle2D.Float(50,50,150,150));
autoShape1.setRoundRadius(autoShape1.getWidth()/3);

//Insert a rectangle with one round corner and set its radius
IAutoShape autoShape2=iSlide.getShapes().appendShape(ShapeType.ONE_ROUND_CORNER_RECTANGLE,new Rectangle2D.Float(250,50,150,150));
autoShape2.setRoundRadius(autoShape2.getWidth()/3);

//Insert a rectangle with one round corner and which one round corner is snipped and set its radius
IAutoShape autoShape3=iSlide.getShapes().appendShape(ShapeType.ONE_SNIP_ONE_ROUND_CORNER_RECTANGLE,new Rectangle2D.Float(450,50,150,150));
autoShape3.setRoundRadius(autoShape3.getWidth()/3);

//Insert a rectangle with two diagonal round corners and set its radius
IAutoShape autoShape4=iSlide.getShapes().appendShape(ShapeType.TWO_DIAGONAL_ROUND_CORNER_RECTANGLE,new Rectangle2D.Float(50,250,150,150));
autoShape4.setRoundRadius(autoShape4.getWidth()/3);

//Insert a rectangle with two same side round corners and set its radius
IAutoShape autoShape5=iSlide.getShapes().appendShape(ShapeType.TWO_SAMESIDE_ROUND_CORNER_RECTANGLE,new Rectangle2D.Float(250,250,150,150));
autoShape5.setRoundRadius(autoShape5.getWidth()/3);
```

---

# Spire.Presentation Rounded Rectangle
## Set radius of rounded rectangle in presentation
```java
//create a PPT document
Presentation presentation = new Presentation();

//insert a rounded rectangle and set its radious
presentation.getSlides().get(0).getShapes().insertRoundRectangle(0, 160, 180, 100, 200, 10);

//append a rounded rectangle and set its radius
IAutoShape shape = presentation.getSlides().get(0).getShapes().appendRoundRectangle(380, 180, 100, 200, 100);

//set the color and fill style of shape
shape.getFill().setFillType( FillFormatType.SOLID);
shape.getFill().getSolidColor().setColor(Color.green);
shape.getShapeStyle().getLineColor().setColor(Color.white);

//rotate the shape to 90 degree
shape.setRotation(90);
```

---

# Spire Presentation Rectangle Formatting
## Set fill and line format for a rectangle shape in PowerPoint presentation
```java
//create a PPT document
Presentation presentation = new Presentation();

//add a shape
Rectangle2D rect = new Rectangle2D.Float((float) presentation.getSlideSize().getSize().getWidth()/ 2 - 100, 100, 200, 100);
IAutoShape shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, rect);

//set the fill format of shape
shape.getFill().setFillType(FillFormatType.SOLID);
shape.getFill().getSolidColor().setColor(Color.BLUE);

//set the fill format of line
shape.getLine().setFillType(FillFormatType.SOLID);
shape.getLine().getSolidFillColor().setColor(Color.DARK_GRAY);
```

---

# Spire.Presentation Shape Shadow Effect
## Apply inner shadow effect to a shape in PowerPoint presentation
```java
//create an instance of presentation document
Presentation ppt = new Presentation();

//get the first slide
ISlide slide = ppt.getSlides().get(0);

//add a shape to slide
Rectangle2D rect1 = new Rectangle2D.Float(200, 150, 300, 120);
IAutoShape shape = slide.getShapes().appendShape(ShapeType.RECTANGLE, rect1);
shape.getFill().setFillType(FillFormatType.SOLID);
shape.getFill().getSolidColor().setColor(Color.LIGHT_GRAY);
shape.getLine().setFillType(FillFormatType.NONE);
shape.getTextFrame().setText("This demo shows how to apply shadow effect to shape.");
shape.getTextFrame().getTextRange().getFill().setFillType(FillFormatType.SOLID);
shape.getTextFrame().getTextRange().getFill().getSolidColor().setColor(Color.BLACK);

//create an inner shadow effect through InnerShadowEffect object
InnerShadowEffect innerShadow = new InnerShadowEffect();
innerShadow.setBlurRadius(20);
innerShadow.setDirection(0);
innerShadow.setDistance(0);
innerShadow.getColorFormat().setColor(Color.BLACK);

//apply the shadow effect to shape
shape.getEffectDag().setInnerShadowEffect(innerShadow);
```

---

# Spire Presentation Shape to Image Converter
## Convert shapes in a PowerPoint presentation to images
```java
// Create a PPT document
Presentation ppt = new Presentation();
ppt.loadFromFile(input);

// Iterate through shapes and convert to images
for (int i = 0; i < ppt.getSlides().get(0).getShapes().getCount(); i++)
{
    // Extract shape as image
    BufferedImage image = ppt.getSlides().get(0).getShapes().saveAsImage(i);
}
```

---

# Spire Presentation Shape to Image Conversion
## Convert presentation shapes to images with custom resolution
```java
// Load presentation
Presentation ppt = new Presentation();
ppt.loadFromFile(input);

// Get the first slide
ISlide slide = ppt.getSlides().get(0);

// Convert each shape to image with custom resolution
for (int i = 0; i < slide.getShapes().getCount(); i++){
    // Save shape as image with custom resolution (300x300)
    BufferedImage image = slide.getShapes().saveAsImage(i, 300, 300);
    ImageIO.write(image, "PNG", new File(fileName));
}

// Dispose presentation
ppt.dispose();
```

---

# Spire Presentation Custom Path Animation
## Add custom path animation to a shape in PowerPoint
```java
//Create a PowerPoint Document
Presentation ppt = new Presentation();

//Add shape
IAutoShape shape = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle(0, 0, 200, 200));

//Add animation
AnimationEffect effect = ppt.getSlides().get(0).getTimeline().getMainSequence().addEffect(shape, AnimationEffectType.PATH_USER);
CommonBehaviorCollection common = effect.getCommonBehaviorCollection();
AnimationMotion motion = (AnimationMotion) common.get(0);
motion.setOrigin(AnimationMotionOrigin.LAYOUT);
motion.setPathEditMode(AnimationMotionPathEditMode.RELATIVE);

//add MotionPath
MotionPath motionPath = new MotionPath();
motionPath.addPathPoints(MotionCommandPathType.MOVE_TO, new Point2D.Float[]{new Point2D.Float(0, 0)}, MotionPathPointsType.CURVE_AUTO, true);
motionPath.addPathPoints(MotionCommandPathType.LINE_TO, new Point2D.Float[]{new Point2D.Float(0.1f, 0.1f)}, MotionPathPointsType.CURVE_AUTO, true);
motionPath.addPathPoints(MotionCommandPathType.LINE_TO, new Point2D.Float[]{new Point2D.Float(-0.1f, 0.2f)}, MotionPathPointsType.CURVE_AUTO, true);
motionPath.addPathPoints(MotionCommandPathType.END, new Point2D.Float[]{}, MotionPathPointsType.CURVE_AUTO, true);
motion.setPath(motionPath);
```

---

# Spire.Presentation Exit Animation
## Adding exit animation to a shape in a PowerPoint presentation
```java
//Create PPT document
Presentation presentation = new Presentation();

//Get the first slide
ISlide slide = presentation.getSlides().get(0);

//Add a shape to the slide
IShape shape = slide.getShapes().appendShape(ShapeType.FIVE_POINTED_STAR, new Rectangle2D.Double(250, 100, 200, 200));

//Add animation effect to the shape
AnimationEffect effect = slide.getTimeline().getMainSequence().addEffect(shape, AnimationEffectType.RANDOM_BARS);

//Change effect type from entrance to exit
effect.setPresetClassType(TimeNodePresetClassType.EXIT);
```

---

# Presentation Shape Animations
## Add animations to slides and shapes in a presentation
```java
//Set the animation of slide to Circle
presentation.getSlides().get(0).getSlideShowTransition().setType(TransitionType.CIRCLE);

//Append new shape - Triangle
IAutoShape shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.TRIANGLE, new Rectangle2D.Double(100, 280, 80, 80));
//Set the color of shape
shape.getFill().setFillType(FillFormatType.SOLID);
shape.getFill().getSolidColor().setColor(Color.orange);
shape.getShapeStyle().getLineColor().setColor(Color.white);
//Set the animation of shape
shape.getSlide().getTimeline().getMainSequence().addEffect(shape, AnimationEffectType.PATH_4_POINT_STAR);

//Append new shape - Rectangle and set animation
shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle2D.Double(210, 280, 150, 80));
shape.getFill().setFillType(FillFormatType.SOLID);
shape.getFill().getSolidColor().setColor(Color.orange);
shape.getShapeStyle().getLineColor().setColor(Color.white);
shape.appendTextFrame("Animated Shape");
shape.getSlide().getTimeline().getMainSequence().addEffect(shape, AnimationEffectType.FADED_SWIVEL);

//Append new shape - Cloud and set the animation
shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.CLOUD, new Rectangle2D.Double(390, 280, 80, 80));
shape.getFill().setFillType(FillFormatType.SOLID);
shape.getFill().getSolidColor().setColor(Color.white);
shape.getShapeStyle().getLineColor().setColor(Color.orange);
shape.getSlide().getTimeline().getMainSequence().addEffect(shape, AnimationEffectType.FADED_ZOOM);
```

---

# Apply Animation on Chart
## This code demonstrates how to apply animation effects to charts in presentations
```java
//Get first shape in first slide
IShape shape=ppt.getSlides().get(0).getShapes().get(0);
if(shape instanceof IChart)
{
    IChart chart = (IChart) shape;
    //Apply Fly animation effect on the chart
    AnimationEffect effect=ppt.getSlides().get(0).getTimeline().getMainSequence().addEffect(chart,AnimationEffectType.FLY);
    //Set the BuildType as SERIES
    effect.getGraphicAnimation().setBuildType(GraphicBuildType.BUILD_AS_SERIES);
}
```

---

# Apply Animation to PowerPoint Shape
## Demonstrates how to apply animation effects to shapes in a PowerPoint presentation
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Get the first slide
ISlide slide = ppt.getSlides().get(0);

//Insert a rectangle in the slide and fill the shape
IAutoShape shape = slide.getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle2D.Double(100, 150, 200, 80));
shape.getFill().setFillType(FillFormatType.SOLID);
shape.getFill().getSolidColor().setColor(Color.gray);
shape.getShapeStyle().getLineColor().setColor(Color.white);
shape.appendTextFrame("Animated Shape");

//Apply FadedSwivel animation effect to the shape
shape.getSlide().getTimeline().getMainSequence().addEffect(shape, AnimationEffectType.FADED_SWIVEL);
```

---

# Spire.Presentation Text Animation
## Apply animation effect to text in PowerPoint presentation
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Get the first slide
ISlide slide = ppt.getSlides().get(0);

//Add a shape to the slide
IAutoShape shape = slide.getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle2D.Double(250, 150, 200, 100));
shape.getFill().setFillType(FillFormatType.SOLID);
shape.getFill().getSolidColor().setColor(Color.gray);
shape.getShapeStyle().getLineColor().setColor(Color.white);
shape.appendTextFrame("This demo shows how to apply animation on text in PPT document.");

//Apply animation to the text in shape
AnimationEffect animation = shape.getSlide().getTimeline().getMainSequence().addEffect(shape, AnimationEffectType.FLOAT);
animation.setStartEndParagraphs(0, 0);
```

---

# Spire Presentation Animation Effect Information
## Extract animation effect details from PowerPoint slides
```java
//Create a PPT document
Presentation presentation = new Presentation();

//Load the document from disk
presentation.loadFromFile("data/Animation.pptx");

StringBuilder sb = new StringBuilder();
//Travel each slide
for(Object slideobj:presentation.getSlides()){
    ISlide slide = (ISlide)slideobj;
    //Travel all animation effects in a slide
    for(Object effectobj:slide.getTimeline().getMainSequence() ){
        AnimationEffect effect = (AnimationEffect)effectobj;

        //Get the animation effect type
        AnimationEffectType animationEffectType = effect.getAnimationEffectType();
        sb.append("animation effect type:"+animationEffectType+"\n");

        //Get the slide number where the animation is located
        int slideNumber = slide.getSlideNumber();
        sb.append("page number:"+slideNumber+"\n");

        //Get shape name
        String shapeName = effect.getShapeTarget().getName();
        sb.append("shape name:"+shapeName+"\n"+"\n");
    }
}
```

---

# spire.presentation animation
## set animation for animate text
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Set the AnimateType as Letter
ppt.getSlides().get(0).getTimeline().getMainSequence().get(0).setIterateType(AnimateType.Letter);

//Set the IterateTimeValue for the animate text
ppt.getSlides().get(0).getTimeline().getMainSequence().get(0).setIterateTimeValue(10);
```

---

# spire presentation animation repeat type
## set animation repeat type for presentation slides
```java
//Get the first slide
ISlide slide = presentation.getSlides().get(0);
AnimationEffectCollection animations = slide.getTimeline().getMainSequence();
animations.get(0).getTiming().setAnimationRepeatType(AnimationRepeatType.UtilEndOfSlide);
```

---

# Spire.Presentation Get Embedded Fonts
## Extract embedded fonts from a PowerPoint presentation

// Load a PowerPoint presentation
Presentation ppt = new Presentation();
ppt.loadFromFile(inputFile);

// Get embedded fonts from the presentation
ArrayList<String> embedFonts = ppt.getEmbedFonts();

// Process each embedded font
for(int i=0; i<embedFonts.size();i++)
{
    // Process each font (e.g., display, analyze, etc.)
    String font = embedFonts.get(i);
}

// Dispose of the Presentation object
ppt.dispose();

---

# Spire.Presentation Load Encrypted Stream
## Load an encrypted PowerPoint presentation using a password
```java
FileInputStream fis = new FileInputStream("data/OpenEncryptedPPT.pptx");

// Create a Presentation instance
Presentation ppt = new Presentation();

// Specify the password for decryption
String password = "123456";

// Load the encrypted stream with the provided password
ppt.loadFromStream(fis, FileFormat.AUTO, password);
```

---

# Load PowerPoint from Stream
## This example demonstrates how to load a PowerPoint presentation from a stream using Spire.Presentation for Java.

```java
//create an instance of presentation document
Presentation ppt = new Presentation();

//load PowerPoint file from stream
File file = new File(input);
FileInputStream in = new FileInputStream(file);
ppt.loadFromStream(in, FileFormat.PPTX_2013);
```

---

# Spire Presentation Loop Configuration
## Configure PowerPoint presentation to loop continuously with animations and narration
```java
//create an instance of presentation document
Presentation ppt = new Presentation();

//set the Boolean value of ShowLoop as true
ppt.setShowLoop(true);

//set the PowerPoint document to show animation and narration
ppt.setShowAnimation(true);
ppt.setShowNarration(true);

//use slide transition timings to advance slide
ppt.setUseTimings(true);
```

---

# Presentation Page Setup
## Set up slide size, orientation, and type in a PowerPoint presentation
```java
//create PPT document
Presentation ppt = new Presentation();

//set the size of slides
ppt.setPageSize(600,600,false);
ppt.getSlideSize().setOrientation(SlideOrienation.PORTRAIT);
ppt.getSlideSize().setType(SlideSizeType.CUSTOM);
```

---

# Spire.Presentation Save to Stream
## Save PowerPoint presentation to output stream
```java
// Create PowerPoint file
Presentation presentation = new Presentation();

// Save to Stream
File outFile = new File("output/saveToStream.pptx");
OutputStream outputStream = new FileOutputStream(outFile);
presentation.saveToFile(outputStream, FileFormat.PPTX_2013);
```

---

# Spire Presentation Kiosk Mode
## Set presentation show type as kiosk
```java
//create an instance of presentation document
Presentation ppt = new Presentation();

//specify the presentation show type as kiosk
ppt.setShowType(SlideShowType.Kiosk);
```

---

# Spire Presentation Split PPT
## Split PowerPoint presentation into individual slides
```java
//create an instance of presentation document
Presentation ppt = new Presentation();

//load file
ppt.loadFromFile(input);

for (int i = 0; i < ppt.getSlides().getCount(); i++)
{
    //initialize another instance of Presentation, and remove the blank slide
    Presentation newppt = new Presentation();
    newppt.getSlides().removeAt(0);

    //append the specified slide from old presentation to the new one
    newppt.getSlides().append(ppt.getSlides().get(i));

    //save the document
    String result = outputPath + String.format("SplitPPT-%d.pptx", i);
    newppt.saveToFile(result, FileFormat.PPTX_2013);
}
```

---

# Presentation Built-in Properties Extraction
## Retrieve built-in document properties from a PowerPoint presentation
```java
//create PPT document
Presentation presentation = new Presentation();

//get the builtin properties
String application = presentation.getDocumentProperty().getApplication();
String author = presentation.getDocumentProperty().getAuthor();
String company = presentation.getDocumentProperty().getCompany();
String keywords = presentation.getDocumentProperty().getKeywords();
String comments = presentation.getDocumentProperty().getComments();
String category = presentation.getDocumentProperty().getCategory();
String title = presentation.getDocumentProperty().getTitle();
String subject = presentation.getDocumentProperty().getSubject();

//Create StringBuilder to save
StringBuilder content = new StringBuilder();
content.append("DocumentProperty.Application: " + application);
content.append("\r\nDocumentProperty.Author: " + author);
content.append("\r\nDocumentProperty.Company " + company);
content.append("\r\nDocumentProperty.Keywords: " + keywords);
content.append("\r\nDocumentProperty.Comments: " + comments);
content.append("\r\nDocumentProperty.Category: " + category);
content.append("\r\nDocumentProperty.Title: " + title);
content.append("\r\nDocumentProperty.Subject: " + subject);
```

---

# Mark Presentation as Final
## Set the MarkAsFinal property for a PowerPoint presentation
```java
//create PPT document
Presentation presentation = new Presentation();

//mark the document as final
presentation.getDocumentProperty().set("_MarkAsFinal", true);
```

---

# Spire.Presentation Document Properties
## Set document properties for a PowerPoint presentation
```java
//create PPT document
Presentation presentation = new Presentation();

//set the DocumentProperty of PPT document
presentation.getDocumentProperty().setApplication("Spire.Presentation");
presentation.getDocumentProperty().setAuthor("E-iceblue");
presentation.getDocumentProperty().setCompany("E-iceblue Co., Ltd.");
presentation.getDocumentProperty().setKeywords("Demo File");
presentation.getDocumentProperty().setComments("This file is used to test Spire.Presentation.");
presentation.getDocumentProperty().setCategory("Demo");
presentation.getDocumentProperty().setTitle("This is a demo file.");
presentation.getDocumentProperty().setSubject("Test");
```

---

# Spire.Presentation Document Properties
## Set properties for presentation template
```java
//create a document
Presentation presentation = new Presentation();

//set the DocumentProperty
presentation.getDocumentProperty().setApplication("Spire.Presentation");
presentation.getDocumentProperty().setAuthor("E-iceblue");
presentation.getDocumentProperty().setCompany("E-iceblue Co., Ltd.");
presentation.getDocumentProperty().setKeywords("Demo File");
presentation.getDocumentProperty().setComments("This file is used to test Spire.Presentation.");
presentation.getDocumentProperty().setCategory("Demo");
presentation.getDocumentProperty().setTitle("This is a demo file.");
presentation.getDocumentProperty().setSubject("Test");
```

---

# Add Digital Signature to PowerPoint
## Add digital signature to presentation using certificate file
```java
//Create a presentation
Presentation ppt = new Presentation();

//Add a digital signature
ppt.addDigitalSignature("data/gary.pfx", "e-iceblue", "Gary", new Date());
```

---

# Spire.Presentation Digital Signature Check
## Check if a PowerPoint document has a digital signature
```java
Presentation ppt = new Presentation();
ppt.loadFromFile("data/digitalSignature.pptx");

//Check if the file contains a digital signature
boolean isSigned = ppt.isDigitallySigned();
```

---

# Spire Presentation Password Protection Check
## Check if a PowerPoint presentation is password protected
```java
String input="data/template_Ppt_4.pptx";

//Create a PPT document
Presentation ppt = new Presentation();

//Check whether a PPT document is protected with password 
boolean password =ppt.isPasswordProtected(input);
```

---

# Presentation Encryption
## Encrypt a PowerPoint presentation with a password
```java
// Create PPT document
Presentation presentation = new Presentation();

// Load the PPT document
presentation.loadFromFile(input);
String strPassword = "e-iceblue";

// Encrypt the document with the password
presentation.encrypt(strPassword);
```

---

# Spire Presentation Password Modification
## Modify password of an encrypted PowerPoint presentation
```java
//create a PowerPoint document
Presentation presentation = new Presentation();

//load the file from disk with original password
presentation.loadFromFile(input, "123456");

//remove the encryption
presentation.removeEncryption();

//protect the document by setting a new password
presentation.protect("654321");
```

---

# Spire.Presentation encrypted file handling
## Open and save encrypted PowerPoint presentations
```java
//create a PowerPoint document
Presentation presentation = new Presentation();

//load the file from disk with original password
presentation.loadFromFile(input, "123456");

//save as a new PPT with original password
presentation.saveToFile(output, FileFormat.PPTX_2013);
presentation.dispose();
```

---

# Remove Digital Signature from PowerPoint
## This code demonstrates how to remove digital signatures from a PowerPoint presentation
```java
public class removeDigitalSignature {
    public static void main(String[] args) throws Exception {
        // Create a presentation object
        Presentation ppt = new Presentation();

        // If the file contains the digital signature
        if (ppt.isDigitallySigned()) {
            // Removes digital signature
            ppt.removeAllDigitalSignatures();
        }
    }
}
```

---

# Remove PowerPoint Encryption
## Remove encryption protection from PowerPoint presentation
```java
// Create a PowerPoint document
Presentation presentation = new Presentation();

// Remove encryption
presentation.removeEncryption();
```

---

# Spire.Presentation Document Security
## Set document to read-only with password protection
```java
// Create a PowerPoint document
Presentation presentation = new Presentation();

// Protect the document with the password
String password = "123456";
presentation.protect(password);
```

---

# spire.presentation background
## Set slide background with image
```java
//Set the background of the first slide to image
presentation.getSlides().get(0).getSlideBackground().setType(BackgroundType.CUSTOM);
presentation.getSlides().get(0).getSlideBackground().getFill().setFillType(FillFormatType.PICTURE);
presentation.getSlides().get(0).getSlideBackground().getFill().getPictureFill().setAlignment(RectangleAlignment.NONE);
presentation.getSlides().get(0).getSlideBackground().getFill().getPictureFill().setFillType(PictureFillType.TILE);
presentation.getSlides().get(0).getSlideBackground().getFill().getPictureFill().getPicture().setUrl((new java.io.File("image_path")).getAbsolutePath());
```

---

# Spire.Presentation Gradient Background
## Set gradient background for presentation slide
```java
//get the first slide
ISlide slide = presentation.getSlides().get(0);

//set the background to gradient
slide.getSlideBackground().setType(BackgroundType.CUSTOM);
slide.getSlideBackground().getFill().setFillType(FillFormatType.GRADIENT);

//add gradient stops
slide.getSlideBackground().getFill().getGradient().getGradientStops().append(0.1f, Color.CYAN);
slide.getSlideBackground().getFill().getGradient().getGradientStops().append(0.7f, Color.LIGHT_GRAY);

//set gradient shape type
slide.getSlideBackground().getFill().getGradient().setGradientShape(GradientShapeType.LINEAR);

//set the angle
slide.getSlideBackground().getFill().getGradient().getLinearGradientFill().setAngle(45);
```

---

# Spire Presentation Master Background
## Set custom solid color background for presentation master slide
```java
//set the slide background of master
presentation.getMasters().get(0).getSlideBackground().setType(BackgroundType.CUSTOM);
presentation.getMasters().get(0).getSlideBackground().getFill().setFillType(FillFormatType.SOLID);
presentation.getMasters().get(0).getSlideBackground().getFill().getSolidColor().setColor(Color.LIGHT_GRAY);
```

---

# Spire Presentation Error Bars Formatting
## Add and format error bars for charts in PowerPoint presentations
```java
//get the column chart on the first slide and set chart title.
IChart columnChart = (IChart)presentation.getSlides().get(0).getShapes().get(0);
columnChart.getChartTitle().getTextProperties().setText("Vertical Error Bars");

//add Y (Vertical) Error Bars.
//get Y error bars of the first chart series.
IErrorBarsFormat errorBarsYFormat1 = columnChart.getSeries().get(0).getErrorBarsYFormat();

//set end cap.
errorBarsYFormat1.setErrorBarNoEndCap(false);

//specify direction.
errorBarsYFormat1.setErrorBarSimType((ErrorBarSimpleType.PLUS).getValue());

//specify error amount type.
errorBarsYFormat1.setErrorBarvType((ErrorValueType.STANDARD_ERROR).getValue());

//set value.
errorBarsYFormat1.setErrorBarVal(0.3f);

//set line format.
errorBarsYFormat1.getLine().setFillType(FillFormatType.SOLID);
errorBarsYFormat1.getLine().getSolidFillColor().setColor(Color.RED);
errorBarsYFormat1.getLine().setWidth(1);

//get the bubble chart on the second slide and set chart title.
IChart bubbleChart = (IChart)presentation.getSlides().get(1).getShapes().get(0);
bubbleChart.getChartTitle().getTextProperties().setText("Vertical and Horizontal Error Bars");

//add X (Horizontal) and Y (Vertical) Error Bars.
//get X error bars of the first chart series.
IErrorBarsFormat errorBarsXFormat = bubbleChart.getSeries().get(0).getErrorBarsXFormat();

//set end cap.
errorBarsXFormat.setErrorBarNoEndCap(false);

//specify direction.
errorBarsXFormat.setErrorBarvType((ErrorBarSimpleType.BOTH).getValue());

//specify error amount type.
errorBarsXFormat.setErrorBarvType((ErrorValueType.STANDARD_ERROR).getValue());

//set value.
errorBarsXFormat.setErrorBarVal(0.3f);

//get Y error bars of the first chart series.
IErrorBarsFormat errorBarsYFormat2 = bubbleChart.getSeries().get(0).getErrorBarsYFormat();

//set end cap.
errorBarsYFormat2.setErrorBarNoEndCap(false);

//specify direction.
errorBarsYFormat2.setErrorBarvType((ErrorBarSimpleType.BOTH).getValue());

//specify error amount type.
errorBarsYFormat2.setErrorBarvType((ErrorValueType.STANDARD_ERROR).getValue());

//set value.
errorBarsYFormat2.setErrorBarVal(0.3f);
```

---

# Spire Presentation Custom Error Bars
## Add custom error bars to a bubble chart in PowerPoint presentation
```java
//get the bubble chart on the first slide
IChart bubbleChart = (IChart)ppt.getSlides().get(0).getShapes().get(0) ;

//get X error bars of the first chart series
IErrorBarsFormat errorBarsXFormat = bubbleChart.getSeries().get(0).getErrorBarsXFormat();

//specify error amount type as custom error bars
errorBarsXFormat.setErrorBarvType((ErrorValueType.CUSTOM_ERROR_BARS).getValue());

//set the minus and plus value of the X error bars
errorBarsXFormat.setMinusVal(0.5f);
errorBarsXFormat.setPlusVal( 0.5f);

//get Y error bars of the first chart series
IErrorBarsFormat errorBarsYFormat = bubbleChart.getSeries().get(0).getErrorBarsYFormat();

//specify error amount type as custom error bars
errorBarsYFormat.setErrorBarvType((ErrorValueType.CUSTOM_ERROR_BARS).getValue());

//set the minus and plus value of the Y error bars
errorBarsYFormat.setMinusVal(1f);
errorBarsYFormat.setPlusVal( 1f);
```

---

# Spire Presentation Secondary Value Axis
## Add a secondary value axis to a chart in PowerPoint presentation
```java
//get the chart from the PowerPoint file.
IChart chart = (IChart)presentation.getSlides().get(0).getShapes().get(0);

//add a secondary axis to display the value of Series 3.
chart.getSeries().get(2).setUseSecondAxis(true);

//set the grid line of secondary axis as invisible.
chart.getSecondaryCategoryAxis().getMajorGridTextLines().setFillType(FillFormatType.NONE);
```

---

# spire presentation chart data label shadow effect
## add shadow effect to data label in presentation chart
```java
//get the chart.
IChart chart =(IChart)presentation.getSlides().get(0).getShapes().get(0);

//add a data label to the first chart series.
ChartDataLabelCollection dataLabels = chart.getSeries().get(0).getDataLabels();
ChartDataLabel Label = dataLabels.add();
Label.setLabelValueVisible(true);

//add outer shadow effect to the data label.
Label.getEffect().setOuterShadowEffect(new OuterShadowEffect());

//set shadow color.
Label.getEffect().getOuterShadowEffect().getColorFormat().setColor( Color.YELLOW);

//set blur.
Label.getEffect().getOuterShadowEffect().setBlurRadius(5);

//set distance.
Label.getEffect().getOuterShadowEffect().setDistance(10);

//set angle.
Label.getEffect().getOuterShadowEffect().setDirection(90f);
```

---

# spire presentation trendline
## add trend line for chart series in PowerPoint
```java
//get the target chart, add trendline for the first data series of the chart and specify the trendline type.
IChart chart = (IChart)presentation.getSlides().get(0).getShapes().get(0);
ITrendlines it = chart.getSeries().get(0).addTrendLine(TrendlineSimpleType.LINEAR);

//set the trendline properties to determine what should be displayed.
it.setdisplayEquation(false);
it.setdisplayRSquaredValue(false);
```

---

# Spire Presentation Pie Chart
## Auto vary colors for pie chart segments
```java
//Create a PPT file
Presentation ppt = new Presentation();

Rectangle2D.Double rect1 = new Rectangle2D.Double(40, 100, 550, 320);
//Add a pie chart
IChart chart = ppt.getSlides().get(0).getShapes().appendChart(ChartType.PIE, rect1, false);
chart.getChartTitle().getTextProperties().setText("Sales by Quarter");
chart.getChartTitle().getTextProperties().isCentered(true);
chart.getChartTitle().setHeight(30);
chart.hasTitle(true);

//Set whether auto vary color, default value is true
chart.getSeries().get(0).isVaryColor(false);

chart.getSeries().get(0).setDistance(15);
```

---

# Spire.Presentation Legend Styling
## Change color and style for chart legend
```java
//get chart on the first slide
IChart Chart =(IChart)ppt.getSlides().get(0).getShapes().get(0);

//change the fill color
Chart.getChartLegend().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().getFill().setFillType(FillFormatType.SOLID);
Chart.getChartLegend().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().getFill().getSolidColor().setColor(Color.BLUE);

//use italic for the paragraph
Chart.getChartLegend().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().isItalic(TriState.TRUE);
```

---

# Spire Presentation Chart Data Table Font
## Change font properties for chart data table in PowerPoint presentation
```java
//get chart on the first slide
IChart Chart =(IChart)ppt.getSlides().get(0).getShapes().get(0);
Chart.hasDataTable( true);

//add a new paragraph in data table
Chart.getChartDataTable().getText().getParagraphs().append(new ParagraphEx());

//change the font size
Chart.getChartDataTable().getText().getParagraphs().get(0).getDefaultCharacterProperties().setFontHeight(15);
```

---

# spire presentation legend font size
## change font size for chart legend
```java
//get chart on the first slide
IChart Chart = (IChart)ppt.getSlides().get(0).getShapes().get(0);

//change legend font size
Chart.getChartLegend().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().setFontHeight(17);
```

---

# Spire.Presentation Chart Series Name
## Change the name of a chart series in a presentation
```java
//get chart on the first slide
IChart Chart =(IChart)ppt.getSlides().get(0).getShapes().get(0);

//get the ranges of series label
CellRanges cr = Chart.getSeries().getSeriesLabel();

//change the value
cr.get(0).setValue("Changed series name");
```

---

# Spire Presentation Chart Text Formatting
## Change text font in chart elements
```java
//set chart title
chart.getChartTitle().getTextProperties().getParagraphs().get(0).setText("Chart Title");

//change the font of title
chart.getChartTitle().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().setLatinFont(new TextFont("Lucida Sans Unicode"));
chart.getChartTitle().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().getFill().getSolidColor().setKnownColor(KnownColors.BLUE);
chart.getChartTitle().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().setFontHeight(30);

//change the font of legend
chart.getChartLegend().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().getFill().getSolidColor().setKnownColor(KnownColors.DARK_GREEN);
chart.getChartLegend().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().setLatinFont(new TextFont("Lucida Sans Unicode"));

//change the font of series
chart.getPrimaryCategoryAxis().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().getFill().getSolidColor().setKnownColor(KnownColors.RED);
chart.getPrimaryCategoryAxis().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().getFill().setFillType(FillFormatType.SOLID);
chart.getPrimaryCategoryAxis().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().setFontHeight(10);
chart.getPrimaryCategoryAxis().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().setLatinFont(new TextFont("Lucida Sans Unicode"));
```

---

# Spire.Presentation Chart Axis Configuration
## Configure primary and secondary axes in a presentation chart
```java
//get the chart
IChart chart = (IChart) ((ppt.getSlides().get(0).getShapes().get(0) instanceof IChart) ? ppt.getSlides().get(0).getShapes().get(0) : null);
chart.getChartTitle().getTextProperties().getParagraphs().get(0).setText("Chart Title");

//add a secondary axis to display the value of Series 3
chart.getSeries().get(2).setUseSecondAxis(true);

//set the grid line of secondary axis as invisible
chart.getSecondaryValueAxis().getMajorGridTextLines().setFillType(FillFormatType.NONE);

//set bounds of axis value. Before we assign values, we must set IsAutoMax and IsAutoMin as false, otherwise MS PowerPoint will automatically set the values.
chart.getPrimaryValueAxis().isAutoMax(false);
chart.getPrimaryValueAxis().isAutoMin(false);
chart.getSecondaryValueAxis().isAutoMax(false);
chart.getSecondaryValueAxis().isAutoMin(false);
chart.getPrimaryValueAxis().setMinValue(0f);
chart.getPrimaryValueAxis().setMaxValue(5.0f);
chart.getSecondaryValueAxis().setMinValue(0f);
chart.getSecondaryValueAxis().setMaxValue(1.0f);

//set axis line format
chart.getPrimaryValueAxis().getMinorGridLines().setFillType(FillFormatType.SOLID);
chart.getSecondaryValueAxis().getMinorGridLines().setFillType(FillFormatType.SOLID);
chart.getPrimaryValueAxis().getMinorGridLines().setWidth(0.1f);
chart.getSecondaryValueAxis().getMinorGridLines().setWidth(0.1f);
chart.getPrimaryValueAxis().getMinorGridLines().getSolidFillColor().setColor(Color.lightGray);
chart.getSecondaryValueAxis().getMinorGridLines().getSolidFillColor().setColor(Color.lightGray);
chart.getPrimaryValueAxis().getMinorGridLines().setDashStyle(LineDashStyleType.DASH);
chart.getSecondaryValueAxis().getMinorGridLines().setDashStyle(LineDashStyleType.DASH);
chart.getPrimaryValueAxis().getMajorGridTextLines().setWidth(0.3f);
chart.getPrimaryValueAxis().getMajorGridTextLines().getSolidFillColor().setColor(Color.blue);
chart.getSecondaryValueAxis().getMajorGridTextLines().setWidth(0.3f);
chart.getSecondaryValueAxis().getMajorGridTextLines().getSolidFillColor().setColor(Color.blue);
```

---

# Spire Presentation Chart Copy
## Copy chart between PowerPoint presentations
```java
// Create a PPT document
Presentation presentation1 = new Presentation();

// Load the file which contains a chart
presentation1.loadFromFile(input1);

// Get the chart that is going to be copied
IChart chart = (IChart)presentation1.getSlides().get(0).getShapes().get(0);

// Load the second PowerPoint document
Presentation presentation2 = new Presentation();
presentation2.loadFromFile(input2);

// Copy chart from the first document to the second document
presentation2.getSlides().append();
presentation2.getSlides().get(1).getShapes().createChart(chart, new Rectangle2D.Double(100, 100, 500, 300), -1);
```

---

# Copy Chart Within PowerPoint Presentation
## Demonstrates how to copy a chart from one slide to another within the same PowerPoint presentation
```java
//Get the chart that will be copied.
IChart chart = (IChart)presentation.getSlides().get(0).getShapes().get(0);

//Copy the chart from the first slide to the specified location of the second slide within the same document.
ISlide slide1 = presentation.getSlides().append();
Rectangle2D.Double rect1 = new Rectangle2D.Double(100, 100, 500, 300);
slide1.getShapes().createChart(chart, rect1, 0);
```

---

# Spire.Presentation 100% Stacked Bar Chart
## Create and configure a 100% stacked bar chart in a PowerPoint presentation
```java
//create a PowerPoint document.
Presentation presentation = new Presentation();

//add a "Bar100PercentStacked" chart to the first slide.
presentation.getSlideSize().setType(SlideSizeType.SCREEN_16_X_9);
Dimension2D slidesize = presentation.getSlideSize().getSize();

//get the first slide
ISlide slide = presentation.getSlides().get(0);

//append a chart.
Rectangle2D rect = new Rectangle2D.Double(20, 20, slidesize.getWidth() - 40, slidesize.getHeight()- 40);
IChart chart = slide.getShapes().appendChart(ChartType.BAR_100_PERCENT_STACKED, rect);

//set the position of category axis.
chart.getPrimaryCategoryAxis().setPosition(AxisPositionType.LEFT);
chart.getSecondaryCategoryAxis().setPosition(AxisPositionType.LEFT);
chart.getPrimaryCategoryAxis().setTickLabelPosition(TickLabelPositionType.TICK_LABEL_POSITION_LOW);

//set the data, font and format for the series of each column.
for (int c = 0; c < chart.getSeries().getCount(); ++c)
{
    chart.getSeries().get(c).getFill().setFillType(FillFormatType.SOLID);
    chart.getSeries().get(c).setInvertIfNegative(false);
    for (int r = 0; r < chart.getCategories().getCount(); ++r)
    {
        ChartDataLabel label = chart.getSeries().get(c).getDataLabels().add();
        label.setLabelValueVisible(true);
        chart.getSeries().get(c).getDataLabels().get(r).hasDataSource(false);
        chart.getSeries().get(c).getDataLabels().get(r).setNumberFormat("0#\\%");
        chart.getSeries().get(c).getDataLabels().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().setFontHeight(12);
    }
}
//set the color of the Series.
chart.getSeries().get(0).getFill().getSolidColor().setColor(Color.YELLOW);
chart.getSeries().get(1).getFill().getSolidColor().setColor(Color.RED);
chart.getSeries().get(2).getFill().getSolidColor().setColor(Color.GREEN);
TextFont font = new TextFont("Tw Cen MT");

//set the font and size for chartlegend.
for (int k = 0; k < chart.getChartLegend().getEntryTextProperties().length; k++)
{
    TextCharacterProperties[] textProperties = chart.getChartLegend().getEntryTextProperties();
    textProperties[k].setLatinFont(font);
    textProperties[k].setFontHeight(20);
}
```

---

# Spire.Presentation Box and Whisker Chart
## Create a Box and Whisker chart in PowerPoint presentation
```java
// Create a PPT file
Presentation ppt = new Presentation();
// Create a Boxandwhisker chart to the first slide
IChart chart = ppt.getSlides().get(0).getShapes().appendChart(ChartType.BOX_AND_WHISKER, new Rectangle2D.Float(50, 50, 500, 400), false);

// Configure series display options
chart.getSeries().get(0).isShowInnerPoints(false);
chart.getSeries().get(0).isShowOutlierPoints(true);
chart.getSeries().get(0).isShowMeanMarkers(true);
chart.getSeries().get(0).isShowMeanLine(true);
chart.getSeries().get(0).setQuartileCalculationType(QuartileCalculation.ExclusiveMedian);
chart.getSeries().get(1).isShowInnerPoints(false);
chart.getSeries().get(1).isShowOutlierPoints(true);
chart.getSeries().get(1).isShowMeanMarkers(true);
chart.getSeries().get(1).isShowMeanLine(true);
chart.getSeries().get(1).setQuartileCalculationType(QuartileCalculation.InclusiveMedian);
chart.getSeries().get(2).isShowInnerPoints(false);
chart.getSeries().get(2).isShowOutlierPoints(true);
chart.getSeries().get(2).isShowMeanMarkers(true);
chart.getSeries().get(2).isShowMeanLine(true);
chart.getSeries().get(2).setQuartileCalculationType(QuartileCalculation.ExclusiveMedian);

// Chart title and legend
chart.hasLegend(true);
chart.getChartTitle().getTextProperties().setText("BoxAndWhisker");
chart.getChartLegend().setPosition(ChartLegendPositionType.TOP);
```

---

# Create Bubble Chart in PowerPoint
## This code demonstrates how to create and configure a bubble chart in a PowerPoint presentation using Spire.Presentation library.
```java
//add bubble chart
Rectangle2D.Double rect1 = new Rectangle2D.Double(90, 100, 550, 320);
IChart chart = null;
chart = presentation.getSlides().get(0).getShapes().appendChart(ChartType.BUBBLE, rect1, false);

//chart title
chart.getChartTitle().getTextProperties().setText("Bubble Chart");
chart.getChartTitle().getTextProperties().isCentered(true);
chart.getChartTitle().setHeight(30);
chart.hasTitle(true);

//set chart data headers
chart.getChartData().get(0, 0).setText("X-Value");
chart.getChartData().get(0, 1).setText("Y-Value");
chart.getChartData().get(0, 2).setText("Size");

//set series label
chart.getSeries().setSeriesLabel(chart.getChartData().get("B1", "B1"));
chart.getSeries().get(0).setXValues(chart.getChartData().get("A2", "A5"));
chart.getSeries().get(0).setYValues(chart.getChartData().get("B2", "B5"));
chart.getSeries().get(0).getBubbles().add(chart.getChartData().get("C2"));
chart.getSeries().get(0).getBubbles().add(chart.getChartData().get("C3"));
chart.getSeries().get(0).getBubbles().add(chart.getChartData().get("C4"));
chart.getSeries().get(0).getBubbles().add(chart.getChartData().get("C5"));
```

---

# Spire Presentation Clustered Column Chart
## Create a clustered column chart in PowerPoint presentation
```java
//create a PPT file
Presentation presentation = new Presentation();

//add clustered column chart
Rectangle2D rect1 = new Rectangle2D.Double(90, 100, 550, 320);
IChart chart = presentation.getSlides().get(0).getShapes().appendChart(ChartType.COLUMN_CLUSTERED, rect1, false);

//chart title
chart.getChartTitle().getTextProperties().setText("Clustered Column Chart");
chart.getChartTitle().getTextProperties().isCentered(true);
chart.getChartTitle().setHeight(30);
chart.hasTitle(true);

//set series text
chart.getChartData().get(0, 1).setText("Series1");
chart.getChartData().get(0, 2).setText("Series2");

//set category text
chart.getChartData().get(1, 0).setText("Category 1");
chart.getChartData().get(2, 0).setText("Category 2");
chart.getChartData().get(3, 0).setText("Category 3");
chart.getChartData().get(4, 0).setText("Category 4");

//set series label
chart.getSeries().setSeriesLabel(chart.getChartData().get("B1", "C1"));

//set category label
chart.getCategories().setCategoryLabels(chart.getChartData().get("A2", "A5"));

//set values for series
chart.getSeries().get(0).setValues(chart.getChartData().get("B2", "B5"));
chart.getSeries().get(1).setValues(chart.getChartData().get("C2", "C5"));
```

---

# spire presentation combination chart
## create a combination chart with column and line series
```java
//create a presentation instance
Presentation presentation = new Presentation();

//insert a column clustered chart
Rectangle2D.Double rect = new Rectangle2D.Double(100, 100, 550, 320);
IChart chart = presentation.getSlides().get(0).getShapes().appendChart(ChartType.COLUMN_CLUSTERED, rect);

//set chart title
chart.getChartTitle().getTextProperties().setText("Monthly Sales Report");
chart.getChartTitle().getTextProperties().isCentered(true);
chart.getChartTitle().setHeight(30);
chart.hasTitle(true);

//create a dataTable
DataTable dataTable = new DataTable();
dataTable.getColumns().add(new DataColumn("Month", DataTypes.DATATABLE_STRING));
dataTable.getColumns().add(new DataColumn("Sales", DataTypes.DATATABLE_INT));
dataTable.getColumns().add(new DataColumn("Growth rate", DataTypes.DATATABLE_DOUBLE));

//import data from dataTable to chart data
for (int c = 0; c < dataTable.getColumns().size(); c++) {
    chart.getChartData().get(0, c).setText(dataTable.getColumns().get(c).getColumnName());
}
for (int r = 0; r < dataTable.getRows().size(); r++) {
    Object[] datas = dataTable.getRows().get(r).getArrayList();
    for (int c = 0; c < datas.length; c++) {
        chart.getChartData().get(r + 1, c).setValue(datas[c]);
    }
}

//set series labels
chart.getSeries().setSeriesLabel(chart.getChartData().get("B1", "C1"));

//set categories labels
chart.getCategories().setCategoryLabels(chart.getChartData().get("A2", "A7"));

//assign data to series values
chart.getSeries().get(0).setValues(chart.getChartData().get("B2", "B7"));
chart.getSeries().get(1).setValues(chart.getChartData().get("C2", "C7"));

//change the chart type of series 2 to line with markers
chart.getSeries().get(1).setType(ChartType.LINE_MARKERS);

//plot data of series 2 on the secondary axis
chart.getSeries().get(1).setUseSecondAxis(true);

//set the number format as percentage
chart.getSecondaryValueAxis().setNumberFormat("0%");

//hide grid links of secondary axis
chart.getSecondaryValueAxis().getMajorGridTextLines().setFillType(FillFormatType.NONE);

//set overlap
chart.setOverLap(-50);

//set gap width
chart.setGapDepth(200);
```

---

# Create 3D Cylinder Clustered Chart
## This code demonstrates how to create a 3D cylinder clustered chart in a presentation using Spire.Presentation for Java.
```java
// Create a presentation instance
Presentation presentation = new Presentation();

// Insert chart
Rectangle2D.Double rect = new Rectangle2D.Double(presentation.getSlideSize().getSize().getWidth() / 2 - 200, 85, 400, 400);
IChart chart = presentation.getSlides().get(0).getShapes().appendChart(ChartType.CYLINDER_3_D_CLUSTERED, rect);

// Add chart Title
chart.getChartTitle().getTextProperties().setText("Report");
chart.getChartTitle().getTextProperties().isCentered(true);
chart.getChartTitle().setHeight(30);
chart.hasTitle(true);

// Set series formatting
chart.getSeries().get(0).getFill().setFillType(FillFormatType.SOLID);
chart.getSeries().get(0).getFill().getSolidColor().setKnownColor(KnownColors.BROWN);
chart.getSeries().get(1).getFill().setFillType(FillFormatType.SOLID);
chart.getSeries().get(1).getFill().getSolidColor().setKnownColor(KnownColors.GREEN);
chart.getSeries().get(2).getFill().setFillType(FillFormatType.SOLID);
chart.getSeries().get(2).getFill().getSolidColor().setKnownColor(KnownColors.ORANGE);

// Set the 3D rotation
chart.getRotationThreeD().setXDegree(10);
chart.getRotationThreeD().setYDegree(10);
```

---

# Creating a Doughnut Chart in PowerPoint
## This code demonstrates how to create and customize a doughnut chart in a PowerPoint presentation using Spire.Presentation for Java
```java
//create a presentation instance
Presentation presentation = new Presentation();

//add a Doughnut chart
Rectangle2D.Double rect = new Rectangle2D.Double(80, 100, 550, 320);
IChart chart = presentation.getSlides().get(0).getShapes().appendChart(ChartType.DOUGHNUT, rect, false);
chart.getChartTitle().getTextProperties().setText("Market share by country");
chart.getChartTitle().getTextProperties().isCentered(true);
chart.getChartTitle().setHeight(30);
chart.getChartData().get(0, 0).setText("Countries");
chart.getChartData().get(0, 1).setText("Sales");
chart.getSeries().setSeriesLabel(chart.getChartData().get("B1", "B1"));
chart.getCategories().setCategoryLabels(chart.getChartData().get("A2", "A5"));
chart.getSeries().get(0).setValues(chart.getChartData().get("B2", "B5"));
for (int i = 0; i < chart.getSeries().get(0).getValues().getCount(); i++) {
    ChartDataPoint cdp = new ChartDataPoint(chart.getSeries().get(0));
    cdp.setIndex(i);
    chart.getSeries().get(0).getDataPoints().add(cdp);
}
//set the series color
chart.getSeries().get(0).getDataPoints().get(0).getFill().setFillType(FillFormatType.SOLID);
chart.getSeries().get(0).getDataPoints().get(0).getFill().getSolidColor().setColor(Color.green);
chart.getSeries().get(0).getDataPoints().get(1).getFill().setFillType(FillFormatType.SOLID);
chart.getSeries().get(0).getDataPoints().get(1).getFill().getSolidColor().setColor(Color.pink);
chart.getSeries().get(0).getDataPoints().get(2).getFill().setFillType(FillFormatType.SOLID);
chart.getSeries().get(0).getDataPoints().get(2).getFill().getSolidColor().setColor(Color.gray);
chart.getSeries().get(0).getDataPoints().get(3).getFill().setFillType(FillFormatType.SOLID);
chart.getSeries().get(0).getDataPoints().get(3).getFill().getSolidColor().setColor(Color.orange);
chart.getSeries().get(0).getDataLabels().setLabelValueVisible(true);
chart.getSeries().get(0).getDataLabels().setPercentValueVisible(true);
```

---

# Creating Histogram Chart in PowerPoint
## This code demonstrates how to create a histogram chart in a PowerPoint presentation using Spire.Presentation for Java.
```java
public class createHistogramChart {
    public static void main(String[] args) throws Exception {
        // Create PPT document
        Presentation ppt = new Presentation();
        
        // Create a Histogram chart to the first slide
        IChart chart = ppt.getSlides().get(0).getShapes().appendChart(ChartType.HISTOGRAM, new Rectangle2D.Float(50, 50, 500, 400), false);
        
        // Set series text
        chart.getChartData().get(0,0).setText("Series 1");
        
        // Set series label
        chart.getSeries().setSeriesLabel(chart.getChartData().get(0,0,0,0));
        
        // Configure chart
        chart.getPrimaryCategoryAxis().setNumberOfBins(7);
        chart.getPrimaryCategoryAxis().setGapWidth(20);
        
        // Chart title
        chart.getChartTitle().getTextProperties().setText("Histogram");
        chart.getChartLegend().setPosition(ChartLegendPositionType.BOTTOM);
    }
}
```

---

# Spire Presentation Line Markers Chart
## Create a line markers chart in PowerPoint presentation
```java
//create a PPT file
Presentation presentation = new Presentation();

//add line markers chart
Rectangle2D rect1 = new Rectangle2D.Double(90, 100, 550, 320);
IChart chart = presentation.getSlides().get(0).getShapes().appendChart(ChartType.LINE_MARKERS, rect1, false);

//chart title
chart.getChartTitle().getTextProperties().setText("Line Makers Chart");
chart.getChartTitle().getTextProperties().isCentered(true);
chart.getChartTitle().setHeight(30);
chart.hasTitle(true);

//set series label
chart.getSeries().setSeriesLabel(chart.getChartData().get("B1", "C1"));
//set category label
chart.getCategories().setCategoryLabels(chart.getChartData().get("A2", "A5"));

//set values for series
chart.getSeries().get(0).setValues(chart.getChartData().get("B2", "B5"));
chart.getSeries().get(1).setValues(chart.getChartData().get("C2", "C5"));
```

---

# Spire Presentation Pareto Chart
## Create a Pareto chart in PowerPoint presentation
```java
//Create a Pareto chart in first slide
IChart chart = ppt.getSlides().get(0).getShapes().appendChart(ChartType.PARETO, new Rectangle2D.Float(50, 50, 500, 400), false);

//Set series label
chart.getSeries().setSeriesLabel(chart.getChartData().get(0,1,0,1));
//Set category label
chart.getCategories().setCategoryLabels(chart.getChartData().get(1,0, 28, 0));
//Set values for series
chart.getSeries().get(0).setValues(chart.getChartData().get(1,1, 28, 1));

chart.getPrimaryCategoryAxis().isBinningByCategory(true);
chart.getSeries().get(1).getLine().getFillFormat().setFillType(FillFormatType.SOLID);
chart.getSeries().get(1).getLine().getFillFormat().getSolidFillColor().setColor(Color.red);
//Chart title
chart.getChartTitle().getTextProperties().setText( "Pareto");
chart.hasLegend(true);
chart.getChartLegend().setPosition(ChartLegendPositionType.BOTTOM);
```

---

# Spire Presentation Pie Chart
## Create a pie chart in PowerPoint presentation with customized title, colors, and data labels

```java
//insert a Pie chart to the first slide and set the chart title.
Rectangle2D rect1 = new Rectangle2D.Double(40, 100, 550, 320);
IChart chart = presentation.getSlides().get(0).getShapes().appendChart(ChartType.PIE, rect1, false);
chart.getChartTitle().getTextProperties().setText("Sales by Quarter");
chart.getChartTitle().getTextProperties().isCentered(true);
chart.getChartTitle().setHeight(30);
chart.hasTitle(true);

//set category labels, series label and series data.
chart.getSeries().setSeriesLabel(chart.getChartData().get("B1", "B1"));
chart.getCategories().setCategoryLabels(chart.getChartData().get("A2", "A5"));
chart.getSeries().get(0).setValues(chart.getChartData().get("B2", "B5"));

//add data points to series and fill each data point with different color.
for (int i = 0; i < chart.getSeries().get(0).getValues().getCount(); i++)
{
    ChartDataPoint cdp = new ChartDataPoint(chart.getSeries().get(0));
    cdp.setIndex(i);
    chart.getSeries().get(0).getDataPoints().add(cdp);
}
chart.getSeries().get(0).getDataPoints().get(0).getFill().setFillType( FillFormatType.SOLID);
chart.getSeries().get(0).getDataPoints().get(0).getFill().getSolidColor().setColor(Color.GREEN);
chart.getSeries().get(0).getDataPoints().get(1).getFill().setFillType( FillFormatType.SOLID);
chart.getSeries().get(0).getDataPoints().get(1).getFill().getSolidColor().setColor(Color.BLUE);
chart.getSeries().get(0).getDataPoints().get(2).getFill().setFillType( FillFormatType.SOLID);
chart.getSeries().get(0).getDataPoints().get(2).getFill().getSolidColor().setColor(Color.PINK);
chart.getSeries().get(0).getDataPoints().get(3).getFill().setFillType( FillFormatType.SOLID);
chart.getSeries().get(0).getDataPoints().get(3).getFill().getSolidColor().setColor(Color.YELLOW);

//set the data labels to display label value and percentage value.
chart.getSeries().get(0).getDataLabels().setLabelValueVisible(true);
chart.getSeries().get(0).getDataLabels().setPercentValueVisible(true);
```

---

# spire.presentation scatter chart
## create scatter chart with markers in presentation
```java
//insert a chart and set chart title and chart type
Rectangle2D.Double rect1 = new  Rectangle2D.Double(90, 100, 550, 320);
IChart chart =  presentation.getSlides().get(0).getShapes().appendChart(ChartType.SCATTER_MARKERS, rect1, false);
chart.getChartTitle().getTextProperties().setText("ScatterMarker Chart");
chart.getChartTitle().getTextProperties().isCentered(true);
chart.getChartTitle().setHeight(30);
chart.hasTitle(true);

//set chart data
Double[] xdata = new Double[] { 2.7, 8.9, 10.0, 12.4 };
Double[] ydata = new Double[] { 3.2, 15.3, 6.7, 8.0 };
chart.getChartData().get(0, 0).setText("X-Value");
chart.getChartData().get(0, 1).setText("Y-Value");
for (int i = 0; i < xdata.length; ++i)
{
    chart.getChartData().get(i + 1, 0).setValue(xdata[i]);
    chart.getChartData().get(i + 1, 1).setValue(ydata[i]);
}
//set the series label
chart.getSeries().setSeriesLabel(chart.getChartData().get("B1", "B1"));

//assign data to X axis, Y axis and Bubbles
chart.getSeries().get(0).setXValues(chart.getChartData().get("A2", "A5"));
chart.getSeries().get(0).setYValues(chart.getChartData().get("B2", "B5"));
```

---

# Create SunBurst Chart
## Create a SunBurst chart in PowerPoint presentation
```java
//Create PPT document
Presentation ppt = new Presentation();
//Create a SunBurst chart to the first slide
IChart chart = ppt.getSlides().get(0).getShapes().appendChart(ChartType.SUN_BURST, new Rectangle2D.Float(50, 50, 500, 400), false);
//Chart title
chart.getChartTitle().getTextProperties().setText( "SunBurst");
chart.hasLegend(true);
chart.getChartLegend().setPosition(ChartLegendPositionType.BOTTOM);
```

---

# Spire.Presentation TreeMap Chart
## Create a TreeMap chart in PowerPoint presentation
```java
//Create PPT document
Presentation ppt = new Presentation();
//Create a TreeMap chart to the first slide
IChart chart = ppt.getSlides().get(0).getShapes().appendChart(ChartType.TREE_MAP, new Rectangle2D.Float(50, 50, 500, 400), false);

//Set series text
chart.getChartData().get(0,3).setText("Series 1");

//Set series labels
chart.getSeries().setSeriesLabel(chart.getChartData().get(0,3,0,3));
//Set categories labels
chart.getCategories().setCategoryLabels(chart.getChartData().get(1,0, 15, 2));
//Assign data to series values
chart.getSeries().get(0).setValues(chart.getChartData().get(1,3, 15, 3));

chart.getSeries().get(0).getDataLabels().setCategoryNameVisible(true);
chart.getSeries().get(0).setTreeMapLabelOption(TreeMapLabelOption.banner);
//Chart title
chart.getChartTitle().getTextProperties().setText( "TreeMap");
chart.hasLegend(true);
chart.getChartLegend().setPosition(ChartLegendPositionType.TOP);
```

---

# Spire.Presentation Waterfall Chart Creation
## Create a waterfall chart with custom data points and formatting
```java
//Create PPT document
Presentation ppt = new Presentation();
//Create a WaterFall chart to the first slide
IChart chart = ppt.getSlides().get(0).getShapes().appendChart(ChartType.WATER_FALL, new Rectangle2D.Float(50, 50, 500, 400), false);

//Set series text
chart.getChartData().get(0,1).setText("Series 1");

//Set category text
String[] categories = { "Category 1", "Category 2", "Category 3", "Category 4", "Category 5", "Category 6", "Category 7" };
for (int i = 0; i < categories.length; i++)
{
    chart.getChartData().get(i+1,0).setText(categories[i]);
}

//Fill data for chart
double[] values = { 100, 20, 50, -40, 130, -60, 70 };
for (int i = 0; i < values.length; i++)
{
    chart.getChartData().get(i+1,1).setNumberValue(values[i]);
}

//Set series labels
chart.getSeries().setSeriesLabel(chart.getChartData().get(0,1,0,1));
//Set categories labels
chart.getCategories().setCategoryLabels(chart.getChartData().get(1,0, categories.length, 0));
//Assign data to series values
chart.getSeries().get(0).setValues(chart.getChartData().get(1,1, values.length, 1));

//Operate the third datapoint of first series
ChartDataPoint chartDataPoint = new ChartDataPoint(chart.getSeries().get(0));
chartDataPoint.setIndex(2);
chartDataPoint.setSetAsTotal(true);
chart.getSeries().get(0).getDataPoints().add(chartDataPoint);

//Operate the sixth datapoint of first series
ChartDataPoint chartDataPoint2 = new ChartDataPoint(chart.getSeries().get(0));
chartDataPoint2.setIndex(5);
chartDataPoint2.setSetAsTotal(true);
chart.getSeries().get(0).getDataPoints().add(chartDataPoint2);

chart.getSeries().get(0).isShowConnectorLines(true);
chart.getSeries().get(0).getDataLabels().setLabelValueVisible(true);

//Chart title
chart.getChartTitle().getTextProperties().setText( "WaterFall");
chart.getChartLegend().setPosition(ChartLegendPositionType.RIGHT);
```

---

# Delete Chart Legend Entries
## Delete specific legend entries from a chart in PowerPoint presentation
```java
//Get the chart.
IChart chart = (IChart)presentation.getSlides().get(0).getShapes().get(0);

//Delete the first and the second legend entries from the chart.
chart.getChartLegend().deleteEntry(0);
chart.getChartLegend().deleteEntry(1);
```

---

# Spire Presentation Doughnut Chart Hole Size
## Set the hole size of a doughnut chart in a PowerPoint presentation
```java
//Get the chart on the first slide
IChart chart = (IChart)ppt.getSlides().get(0).getShapes().get(0);

//Set hole size
chart.getSeries().get(0).setDoughnutHoleSize(55);
```

---

# Spire.Presentation Chart Data Editing
## Edit chart data values in a PowerPoint presentation
```java
//Get chart on the first slide
IChart chart = (IChart)ppt.getSlides().get(0).getShapes().get(0);

//Change the value of the second datapoint of the first series
chart.getSeries().get(0).getValues().get(1).setValue(6);
```

---

# Spire Presentation Explode Pie Chart
## Set explosion distance for a pie chart in a presentation
```java
//Get the chart that needs to set the point explosion.
IChart chart = (IChart)presentation.getSlides().get(0).getShapes().get(0);

chart.getSeries().get(0).setDistance(15);
```

---

# Fill Picture in Chart Marker
## Demonstrates how to fill a picture in a chart marker in a PowerPoint presentation
```java
//Get chart on the first slide
IChart chart = (IChart)ppt.getSlides().get(0).getShapes().get(0);

//Create a ChartDataPoint object and specify the index
ChartDataPoint dataPoint = new ChartDataPoint(chart.getSeries().get(0));
dataPoint.setIndex(0);

//Fill picture in marker
dataPoint.getMarkerFill().getFill().setFillType(FillFormatType.PICTURE);
dataPoint.getMarkerFill().getFill().getPictureFill().getPicture().setEmbedImage(imageData);

//Set marker size
dataPoint.setMarkerSize(20);

//Add the data point in series
chart.getSeries().get(0).getDataPoints().add(dataPoint);
```

---

# Spire.Presentation Chart Data Labels Formatting
## Format chart data labels with custom text, position, font, and color
```java
//Get the chart
IChart chart = (IChart) ((ppt.getSlides().get(0).getShapes().get(0) instanceof IChart) ? ppt.getSlides().get(0).getShapes().get(0) : null);
//Get the chart series
ChartSeriesFormatCollection sers = chart.getSeries();

//Initialize four instances of series label and set parameters of each label
ChartDataLabel cd1 = sers.get(0).getDataLabels().add();
cd1.setPercentageVisible(true);
cd1.getTextFrame().setText("Custom Datalabel1");
cd1.getTextFrame().getTextRange().setFontHeight(12);
cd1.getTextFrame().getTextRange().setLatinFont(new TextFont("Lucida Sans Unicode"));
cd1.getTextFrame().getTextRange().getFill().setFillType(FillFormatType.SOLID);
cd1.getTextFrame().getTextRange().getFill().getSolidColor().setColor(Color.green);

ChartDataLabel cd2 = sers.get(0).getDataLabels().add();
cd2.setPosition(ChartDataLabelPosition.INSIDE_END);
cd2.setPercentageVisible(true);
cd2.getTextFrame().setText("Custom Datalabel2");
cd2.getTextFrame().getTextRange().setFontHeight(10);
cd2.getTextFrame().getTextRange().setLatinFont(new TextFont("Arial"));
cd2.getTextFrame().getTextRange().getFill().setFillType(FillFormatType.SOLID);
cd2.getTextFrame().getTextRange().getFill().getSolidColor().setColor(Color.orange);

ChartDataLabel cd3 = sers.get(0).getDataLabels().add();
cd3.setPosition(ChartDataLabelPosition.CENTER);
cd3.setPercentageVisible(true);
cd3.getTextFrame().setText("Custom Datalabel3");
cd3.getTextFrame().getTextRange().setFontHeight(14);
cd3.getTextFrame().getTextRange().setLatinFont(new TextFont("Calibri"));
cd3.getTextFrame().getTextRange().getFill().setFillType(FillFormatType.SOLID);
cd3.getTextFrame().getTextRange().getFill().getSolidColor().setColor(Color.blue);

ChartDataLabel cd4 = sers.get(0).getDataLabels().add();
cd4.setPosition(ChartDataLabelPosition.INSIDE_BASE);
cd4.setPercentageVisible(true);
cd4.getTextFrame().setText("Custom Datalabel4");
cd4.getTextFrame().getTextRange().setFontHeight(12);
cd4.getTextFrame().getTextRange().setLatinFont(new TextFont("Lucida Sans Unicode"));
cd4.getTextFrame().getTextRange().getFill().setFillType(FillFormatType.SOLID);
cd4.getTextFrame().getTextRange().getFill().getSolidColor().setColor(Color.pink);
```

---

# Spire.Presentation Chart Axis Values
## Get values and units from chart axes
```java
//Get chart on the first slide
IChart chart = (IChart)ppt.getSlides().get(0).getShapes().get(0);

//Get unit from primary category axis
float MajorUnit = chart.getPrimaryCategoryAxis().getMajorUnit();
ChartBaseUnitType type = chart.getPrimaryCategoryAxis().getMajorUnitScale();

//Get values from primary value axis
float minValue = chart.getPrimaryValueAxis().getMinValue();
float maxValue = chart.getPrimaryValueAxis().getMaxValue();
```

---

# Spire.Presentation Chart Axis Labels
## Group two-level axis labels in a chart
```java
//Get the chart from the presentation.
IChart chart = (IChart)presentation.getSlides().get(0).getShapes().get(0);

//Get the category axis from the chart.
IChartAxis chartAxis = chart.getPrimaryCategoryAxis();

//Group the axis labels that have the same first-level label.
if (chartAxis.hasMultiLvlLbl())
{
    chartAxis.isMergeSameLabel(true);
}
```

---

# Spire Presentation Chart Axis and Gridline Control
## Hide chart axes and gridlines in a PowerPoint presentation
```java
//Get chart on the first slide
IChart chart = (IChart)ppt.getSlides().get(0).getShapes().get(0);

//Hide axis
chart.getPrimaryCategoryAxis().isVisible(false);
chart.getPrimaryValueAxis().isVisible(false);

//Remove grid line
chart.getPrimaryValueAxis().getMajorGridTextLines().setFillType(FillFormatType.NONE);
```

---

# Spire.Presentation Chart Series Visibility
## Hide or show a series in a PowerPoint chart
```java
//Get the first slide.
ISlide slide = presentation.getSlides().get(0);

//Get the first chart.
IChart chart = (IChart)slide.getShapes().get(0);

//Hide the first series of the chart.
chart.getSeries().get(0).isHidden(true);

//Show the first series of the chart.
//chart.Series[0].IsHidden = false;
```

---

# Spire Presentation Chart Invert If Negative
## Set chart series to invert colors for negative values
```java
//Get chart on the first slide
IChart chart = (IChart)ppt.getSlides().get(0).getShapes().get(0);

//Set invert if negative
chart.getSeries().get(0).setInvertIfNegative(true);
```

---

# Spire.Presentation Chart Category Axis Modification
## Modify chart category axis properties including major unit and scale
```java
//Get chart on the first slide
IChart chart = (IChart)ppt.getSlides().get(0).getShapes().get(0);

//Modify the major unit
chart.getPrimaryCategoryAxis().isAutoMajor(false);
chart.getPrimaryCategoryAxis().setMajorUnit(1);
chart.getPrimaryCategoryAxis().setMajorUnitScale(ChartBaseUnitType.MONTHS);
```

---

# Spire.Presentation Multiple Category Chart
## Create a chart with multiple categories in a PowerPoint presentation
```java
//Create a PPT file
Presentation presentation = new Presentation();

//Add line markers chart
Rectangle2D rect1 = new Rectangle2D.Double(90, 100, 550, 320);
IChart chart = presentation.getSlides().get(0).getShapes().appendChart(ChartType.COLUMN_CLUSTERED, rect1, false);

//Chart title
chart.getChartTitle().getTextProperties().setText("Multiple-Category");
chart.getChartTitle().getTextProperties().isCentered(true);
chart.getChartTitle().setHeight(30);
chart.hasTitle(true);

//Data for series
Double[] Series1 = new Double[] { 7.7, 8.9, 7.0, 6.0,7.0, 8.0 };

//Set series text
chart.getChartData().get(0,2).setText("Series1");

//Set category text
chart.getChartData().get(1,0).setText("Grp 1");
chart.getChartData().get(3,0).setText("Grp 2");
chart.getChartData().get(5,0).setText("Grp 3");

chart.getChartData().get(1,1).setText("A");
chart.getChartData().get(2,1).setText("B");
chart.getChartData().get(3,1).setText("C");
chart.getChartData().get(4,1).setText("D");
chart.getChartData().get(5,1).setText("E");
chart.getChartData().get(6,1).setText("F");

//Fill data for chart
for (int i = 0; i < Series1.length; ++i) {
    chart.getChartData().get(i + 1, 2).setValue(Series1[i]);
}

//Set series label
chart.getSeries().setSeriesLabel(chart.getChartData().get("C1", "C1"));
//Set category label
chart.getCategories().setCategoryLabels(chart.getChartData().get("A2", "B7"));

//Set values for series
chart.getSeries().get(0).setValues(chart.getChartData().get("C2", "C7"));

//Set if the category axis has multiple levels
chart.getPrimaryCategoryAxis().hasMultiLvlLbl(true);
//Merge same label
chart.getPrimaryCategoryAxis().isMergeSameLabel(true);
```

---

# Spire Presentation Chart Protection
## Protect chart data in PowerPoint presentation
```java
//Get the first shape from slide and convert it as IChart
IChart chart = (IChart)presentation.getSlides().get(0).getShapes().get(0);

//Set the Boolean value of IChart.IsDataProtect as true
chart.isDataProtect(true);
```

---

# Remove Chart from PowerPoint Slide
## This code demonstrates how to remove chart shapes from a PowerPoint slide
```java
//Get the first slide from the document.
ISlide slide = presentation.getSlides().get(0);

//Remove chart from the slide.
for (int i = 0; i < slide.getShapes().getCount(); i++) {
    IShape shape =(IShape)slide.getShapes().get(i);
    if (shape instanceof IChart)
    {
        slide.getShapes().remove(shape);
    }
}
```

---

# Spire Presentation Chart Axis Formatting
## Remove tick marks from chart axis and set number format
```java
//Get the chart that need to be adjusted the number format and remove the tick marks.
IChart chart = (IChart)presentation.getSlides().get(0).getShapes().get(0);

//Set percentage number format for the axis value of chart.
chart.getPrimaryValueAxis().setNumberFormat("0#\\%");

//Remove the tick marks for value axis and category axis.
chart.getPrimaryValueAxis().setMajorTickMark(TickMarkType.TICK_MARK_NONE);
chart.getPrimaryValueAxis().setMinorTickMark(TickMarkType.TICK_MARK_NONE);
chart.getPrimaryCategoryAxis().setMajorTickMark(TickMarkType.TICK_MARK_NONE);
chart.getPrimaryCategoryAxis().setMinorTickMark(TickMarkType.TICK_MARK_NONE);
```

---

# Save Chart as Image
## Extract and save a chart from a PowerPoint presentation as an image
```java
//Create a ppt document and load file
Presentation presentation = new Presentation();
presentation.loadFromFile("input.pptx");

//Save chart as image in .png format
BufferedImage image = presentation.getSlides().get(0).getShapes().saveAsImage(0);
ImageIO.write(image,"PNG",new File("output.png"));
```

---

# Spire Presentation Bubble Chart Scaling
## Scale bubble chart size in PowerPoint presentation
```java
//Create a PowerPoint document.
Presentation presentation = new Presentation();

//Get the chart from the first presentation slide.
IChart chart = (IChart)presentation.getSlides().get(0).getShapes().get(0);

//Scale the bubble size, the range value is from 0 to 300.
chart.setBubbleScale(50);
```

---

# spire presentation chart axis position
## set chart axis position in PowerPoint presentation
```java
//Get chart on the first slide
IChart chart = (IChart)ppt.getSlides().get(0).getShapes().get(0);

//Set axis position
chart.getPrimaryValueAxis().setCrossBetweenType(CrossBetweenType.MIDPOINT_OF_CATEGORY.getValue());
```

---

# Spire.Presentation Chart Axis Configuration
## Set chart axis type to date axis with month scale
```java
//Get the chart
IChart chart = (IChart)presentation.getSlides().get(0).getShapes().get(0);

chart.getPrimaryCategoryAxis().setAxisType(AxisType.DateAxis);
chart.getPrimaryCategoryAxis().setMajorUnitScale(ChartBaseUnitType.MONTHS);
```

---

# Spire Presentation Chart Border Corners
## Set chart border corners to right angle or rounded
```java
//Get chart on the first slide
ISlide slide = ppt.getSlides().get(0);
IChart chart = (IChart)slide.getShapes().get(0);

//Set border as solid
chart.getLine().setFillType(FillFormatType.SOLID);

//Set border to right angle, "false" for right angles, "true" for rounded corners
chart.setBorderRoundedCorners(false);
```

---

# Set Data Label Position in Chart
## This code demonstrates how to set the position of a data label in a chart within a PowerPoint presentation
```java
//Get chart on the first slide
IChart chart = (IChart)ppt.getSlides().get(0).getShapes().get(0);

//Add data label
ChartDataLabel label = chart.getSeries().get(0).getDataLabels().add();
//Set the position of the label
label.setX (2f);
label.setY (2f);
```

---

# Spire Presentation Chart Data Point Coloring
## Set custom colors for data points in a PowerPoint chart
```java
//Get the chart
IChart chart = (IChart) ((ppt.getSlides().get(0).getShapes().get(0) instanceof IChart) ? ppt.getSlides().get(0).getShapes().get(0) : null);
chart.getChartTitle().getTextProperties().setText("Chart Title");

//Initialize an instances of dataPoint
ChartDataPoint cdp1 = new ChartDataPoint(chart.getSeries().get(0));
//Specify the data point order
cdp1.setIndex(0);
//Set the color of the data point
cdp1.getFill().setFillType(FillFormatType.SOLID);
cdp1.getFill().getSolidColor().setKnownColor(KnownColors.ORANGE);

//Add the dataPoint to first series
chart.getSeries().get(0).getDataPoints().add(cdp1);

//Set the color for the other three data points
ChartDataPoint cdp2 = new ChartDataPoint(chart.getSeries().get(0));
cdp2.setIndex(1);
cdp2.getFill().setFillType(FillFormatType.SOLID);
cdp2.getFill().getSolidColor().setKnownColor(KnownColors.GOLD);
chart.getSeries().get(0).getDataPoints().add(cdp2);

ChartDataPoint cdp3 = new ChartDataPoint(chart.getSeries().get(0));
cdp3.setIndex(2);
cdp3.getFill().setFillType(FillFormatType.SOLID);
cdp3.getFill().getSolidColor().setKnownColor(KnownColors.MEDIUM_PURPLE);
chart.getSeries().get(0).getDataPoints().add(cdp3);

ChartDataPoint cdp4 = new ChartDataPoint(chart.getSeries().get(0));
cdp4.setIndex(1);
cdp4.getFill().setFillType(FillFormatType.SOLID);
cdp4.getFill().getSolidColor().setKnownColor(KnownColors.FOREST_GREEN);
chart.getSeries().get(0).getDataPoints().add(cdp4);
```

---

# Spire.Presentation Chart Display Unit
## Set display unit for chart value axis in PowerPoint presentation
```java
//Create PPT document
Presentation ppt = new Presentation();

//Get chart on the first slide
IChart chart = (IChart)ppt.getSlides().get(0).getShapes().get(0);

//Set the display unit
chart.getPrimaryValueAxis().setDisplayUnit(ChartDisplayUnitType.HUNDREDS);
```

---

# Spire.Presentation Chart Axis Distance
## Set distance from axis for chart labels in PowerPoint presentation
```java
//create a powerpoint file
Presentation ppt = new Presentation();

//get the first slide
ISlide slide = ppt.getSlides().get(0);

//Append a chart in slide
Rectangle2D rect = new Rectangle2D.Double(50, 50, 400, 400);
IChart chart = slide.getShapes().appendChart(ChartType.COLUMN_CLUSTERED, rect);

//get the PrimaryCategory axis
IChartAxis chartAxis = chart.getPrimaryCategoryAxis();

//set "Distance from axis"
chartAxis.setLabelsDistance(200);
```

---

# spire presentation chart gap width
## set gap width for chart in presentation
```java
//Create PPT document
Presentation ppt = new Presentation();

//Get chart on the first slide
IChart chart = (IChart)ppt.getSlides().get(0).getShapes().get(0);

//Set gap width
chart.setGapWidth(50);
```

---

# spire presentation legend options
## set chart legend position and size
```java
//Get chart on the first slide
IChart chart = (IChart)ppt.getSlides().get(0).getShapes().get(0);

//Set the legend positon
chart.getChartLegend().setLeft(20);
chart.getChartLegend().setTop(20);

//Set the legend size
chart.getChartLegend().setWidth(250);
chart.getChartLegend().setHeight(30);
```

---

# Spire.Presentation Number Format for Axis
## Set number format for chart axis
```java
//Get chart on the first slide
IChart chart = (IChart)ppt.getSlides().get(0).getShapes().get(0);

//Set the number format
chart.getPrimaryCategoryAxis().setNumberFormat("yyyy");
```

---

# Set Percentage for Chart Labels
## Set percentage values for data labels in a PowerPoint chart
```java
//Get the chart on the first slide
IChart chart = (IChart)ppt.getSlides().get(0).getShapes().get(0);

float dataPontPercent = 0f;

for (int i = 0; i < chart.getSeries().size(); i++)
{
    ChartSeriesDataFormat series = chart.getSeries().get(i);
    //Get the total number
    float total = GetTotal(series.getValues());
    for (int j = 0; j < series.getValues().getCount(); j++) {
        //Get the percent
        dataPontPercent = Float.parseFloat(series.getValues().get(j).getText()) / total * 100;
        //Add data labels
        ChartDataLabel label = series.getDataLabels().add();
        label.setLabelValueVisible(true);
        //Set the percent text for the label
        DecimalFormat df1 = new DecimalFormat("##.00%");
        label.getTextFrame().getParagraphs().get(0).setText(df1.format(dataPontPercent/100));
        label.getTextFrame().getParagraphs().get(0).getTextRanges().get(0).setFontHeight(12);
    }
}

private static float GetTotal(CellRanges ranges)
{
    float total = 0;
    for (int i = 0; i < ranges.getCount(); i++)
    {
        total += Float.parseFloat(ranges.get(i).getText());
    }

    return total;
}
```

---

# Spire Presentation Chart Data Labels Position
## Set position and style of chart data labels in PowerPoint presentation
```java
//Add data label to chart and set its id.
ChartDataLabel label1 = chart.getSeries().get(0).getDataLabels().add();
label1.setID(0);

//Set the default position of data label. This position is relative to the data markers.
label1.setPosition(ChartDataLabelPosition.OUTSIDE_END);

//Set custom position of data label. This position is relative to the default position.
label1.setX(0.1f);
label1.setY(-0.1f);

//Set label value visible
label1.setLabelValueVisible(true);

//Set legend key invisible
label1.setLegendKeyVisible(false);

//Set category name invisible
label1.setCategoryNameVisible(false);

//Set series name invisible
label1.setSeriesNameVisible(false);

//Set Percentage invisible
label1.setPercentageVisible(false);

//Set border style and fill style of data label
label1.getLine().setFillType(FillFormatType.SOLID);
label1.getLine().getSolidFillColor().setColor(Color.blue);
label1.getFill().setFillType(FillFormatType.SOLID);
label1.getFill().getSolidColor().setColor(Color.orange);
```

---

# Spire Presentation Chart Title Rotation
## Set rotation angle for chart title in PowerPoint presentation
```java
//Get the chart
IChart chart = (IChart)presentation.getSlides().get(0).getShapes().get(0);

chart.getChartTitle().getTextProperties().setRotationAngle(-30);
```

---

# Spire.Presentation Chart Data Label Rotation
## Set rotation angle for chart data labels
```java
//Get chart on the slide
IChart Chart = (IChart) ppt.getSlides().get(0).getShapes().get(0);

//Set the rotation angle for the data labels of first series
for (int i = 0; i < Chart.getSeries().get(0).getValues().getCount(); i++) {
    ChartDataLabel label = Chart.getSeries().get(0).getDataLabels().add();
    label.setID(i);
    label.setRotationAngle(45);
}
```

---

# Spire Presentation Chart Value Axis Text Rotation
## Set rotation angle for text on the value axis in a chart
```java
//Get chart on the first slide
IChart chart = (IChart)ppt.getSlides().get(0).getShapes().get(0);

//Set the rotation angle for the text on the value axis
chart.getPrimaryValueAxis().setTextRotationAngle(45);
```

---

# Spire.Presentation Chart Series Overlap
## Set the overlap value for chart series in a PowerPoint presentation
```java
//Get chart on the first slide
IChart chart = (IChart)ppt.getSlides().get(0).getShapes().get(0);

//Set overlap
chart.setOverLap(50);
```

---

# Spire.Presentation Chart Marker Styling
## Set size and style for chart markers in PowerPoint presentations
```java
//Get the chart from the presentation.
IChart chart = (IChart)presentation.getSlides().get(0).getShapes().get(0);

for (int i = 0; i < chart.getSeries().get(0).getValues().getCount(); i++) {
    //Create a ChartDataPoint object and specify the index.
    ChartDataPoint dataPoint = new ChartDataPoint(chart.getSeries().get(0));
    dataPoint.setIndex(i);

    //Set the fill color of the data marker.
    dataPoint.getMarkerFill().getFill().setFillType(FillFormatType.SOLID);
    dataPoint.getMarkerFill().getFill().getSolidColor().setColor(Color.yellow);

    //Set the line color of the data marker.
    dataPoint.getMarkerFill().getLine().setFillType(FillFormatType.SOLID);
    dataPoint.getMarkerFill().getLine().getSolidFillColor().setKnownColor(KnownColors.YELLOW_GREEN);

    //Set the size of the data marker.
    dataPoint.setMarkerSize(20);

    //Set the style of the data marker
    dataPoint.setMarkerStyle(ChartMarkerType.DIAMOND);
    chart.getSeries().get(0).getDataPoints().add(dataPoint);
}
```

---

# Setting Chart Plot Area Size
## Set width and height for chart plot area in a PowerPoint presentation
```java
//Get chart on the first slide
IChart chart = (IChart)presentation.getSlides().get(0).getShapes().get(0);

//Set width and height for chart plot area
chart.getPlotArea().setWidth(250);
chart.getPlotArea().setHeight(300);
```

---

# Spire.Presentation Chart Title Font
## Set font properties for chart title in PowerPoint presentation
```java
//Get the chart from the presentation
IChart chart = (IChart)presentation.getSlides().get(0).getShapes().get(0);

//Set the font for the text on chart title area
chart.getChartTitle().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().setLatinFont(new TextFont("Arial Unicode MS"));
chart.getChartTitle().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().getFill().getSolidColor().setKnownColor(KnownColors.BLUE);
chart.getChartTitle().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().setFontHeight(50);
```

---

# Spire.Presentation Chart Text Formatting
## Set text font for chart legend and axis
```java
//Set the font for the text on Chart Legend area.
chart.getChartLegend().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().getFill().getSolidColor().setKnownColor(KnownColors.GREEN) ;
chart.getChartLegend().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().setLatinFont(new TextFont("Arial Unicode MS"));

//Set the font for the text on Chart Axis area.
chart.getPrimaryCategoryAxis().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().getFill().getSolidColor().setKnownColor(KnownColors.RED);
chart.getPrimaryCategoryAxis().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().getFill().setFillType(FillFormatType.SOLID);
chart.getPrimaryCategoryAxis().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().setFontHeight(10);
chart.getPrimaryCategoryAxis().getTextProperties().getParagraphs().get(0).getDefaultCharacterProperties().setLatinFont(new TextFont("Arial Unicode MS"));
```

---

# Spire Presentation Chart Tick Labels
## Set tick mark labels on category axis in PowerPoint chart
```java
//Get the chart from the PowerPoint slide.
IChart chart = (IChart)presentation.getSlides().get(0).getShapes().get(0);

//Rotate tick labels.
chart.getPrimaryCategoryAxis().setTextRotationAngle(45);

//Specify interval between labels.
chart.getPrimaryCategoryAxis().isAutomaticTickLabelSpacing(false);
chart.getPrimaryCategoryAxis().setTickLabelSpacing(2);

//Change position.
chart.getPrimaryCategoryAxis().setTickLabelPosition(TickLabelPositionType.TICK_LABEL_POSITION_HIGH);
```

---

# spire presentation chart labels
## show chart labels in presentation
```java
//Get chart on the first slide
IChart chart = (IChart)presentation.getSlides().get(0).getShapes().get(0);

//Show data labels
chart.getSeries().get(0).getDataLabels().setLabelValueVisible(true);
chart.getSeries().get(0).getDataLabels().setCategoryNameVisible(true);
chart.getSeries().get(0).getDataLabels().setSeriesNameVisible(true);
```

---

# Vary Colors of Same Series Data Markers
## This code demonstrates how to set different colors for data markers within the same series in a PowerPoint chart.
```java
//Get the chart from the presentation.
IChart chart = (IChart)presentation.getSlides().get(0).getShapes().get(0);

//Create a ChartDataPoint object and specify the index.
ChartDataPoint dataPoint = new ChartDataPoint(chart.getSeries().get(0));
dataPoint.setIndex(0);

//Set the fill color of the data marker.
dataPoint.getMarkerFill().getFill().setFillType(FillFormatType.SOLID);
dataPoint.getMarkerFill().getFill().getSolidColor().setColor(Color.RED);

//Set the line color of the data marker.
dataPoint.getMarkerFill().getLine().setFillType(FillFormatType.SOLID);
dataPoint.getMarkerFill().getLine().getSolidFillColor().setColor(Color.RED);

//Add the data point to the point collection of a series.
chart.getSeries().get(0).getDataPoints().add(dataPoint);

dataPoint = new ChartDataPoint(chart.getSeries().get(0));
dataPoint.setIndex(1);
//Set the fill color of the data marker.
dataPoint.getMarkerFill().getFill().setFillType(FillFormatType.SOLID);
dataPoint.getMarkerFill().getFill().getSolidColor().setColor(Color.BLACK);

//Set the line color of the data marker.
dataPoint.getMarkerFill().getLine().setFillType(FillFormatType.SOLID);
dataPoint.getMarkerFill().getLine().getSolidFillColor().setColor(Color.BLACK);
chart.getSeries().get(0).getDataPoints().add(dataPoint);

dataPoint = new ChartDataPoint(chart.getSeries().get(0));
dataPoint.setIndex(2);
//Set the fill color of the data marker.
dataPoint.getMarkerFill().getFill().setFillType(FillFormatType.SOLID);
dataPoint.getMarkerFill().getFill().getSolidColor().setColor(Color.BLUE);

//Set the line color of the data marker.
dataPoint.getMarkerFill().getLine().setFillType(FillFormatType.SOLID);
dataPoint.getMarkerFill().getLine().getSolidFillColor().setColor(Color.BLUE);
chart.getSeries().get(0).getDataPoints().add(dataPoint);
```

---

# Spire.Presentation Slide Conversion
## Convert all presentation slides to SVG format
```java
// Create presentation object
Presentation ppt = new Presentation();

// Convert all slides to SVG
byte[] bytes = ppt.saveToOneSVG();
```

---

# Spire.Presentation file format conversion
## Convert DPS format files to DPT format
```java
//Load Dps file.
Presentation presentation = new Presentation();
presentation.loadFromFile("data/Sample_dps.dps", FileFormat.DPS);

//Convert to Dpt file.
presentation.saveToFile("output/result.dpt", FileFormat.DPT);
```

---

# Convert DPT to DPS
## Convert a DPT presentation file to DPS format
```java
//Load Dpt file.
Presentation presentation = new Presentation();
presentation.loadFromFile("data/Sample_dpt.dpt", FileFormat.DPT);

//Convert to Dps file.
presentation.saveToFile("output/result.dps", FileFormat.DPS);
```

---

# Spire.Presentation ODP to PDF Conversion
## Convert ODP format presentation to PDF format
```java
Presentation presentation = new Presentation();

//Load ODP file from disk
presentation.loadFromFile("data/toPdf.odp", FileFormat.ODP);

String result = "output/ConvertODPtoPDF_result.pdf";

//Save to file.
presentation.saveToFile(result, FileFormat.PDF);
```

---

# Convert PowerPoint to PDF with Default Font
## Set a default font and convert presentation to PDF
```java
//Create a ppt document
Presentation ppt = new Presentation();
ppt.loadFromFile("data/ConvertPdfWithDefaultFont.pptx");

//The font is preferred to convert to pdf or pictures, when the font used in the document is not installed in the system
Presentation.setDefaultFontName("Arial");

//Save to file
ppt.saveToFile("ConvertPdfWithDefaultFont.pdf", FileFormat.PDF);
```

---

# Spire Presentation Conversion
## Convert PPS to PPTX format
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();
//Load file
ppt.loadFromFile("data/Conversion.pps");

//Save the PPS document to PPTX file format
String result = "output/convertPPSToPPTX_result.pptx";
ppt.saveToFile(result, FileFormat.PPTX_2013);
```

---

# spire presentation to svg conversion
## Convert presentation slides to SVG format
```java
// Create presentation
Presentation ppt = new Presentation();

// Convert specific slides to SVG format
// startSlide: Start slide index, endSlide: End slide index
byte[] bytes = ppt.saveToOneSVG(0, 1);
```

---

# Spire Presentation Convert Individual Slide to HTML
## Convert a single slide from PowerPoint to HTML format
```java
//Create PPT document
Presentation presentation = new Presentation();

//Load the PPT document from disk.
presentation.loadFromFile("data/changeSlidePosition.pptx");

//Get the first slide
ISlide slide = presentation.getSlides().get(0);

//Save the first slide to HTML
slide.SaveToFile("output/individualSlideToHtml_result.html", FileFormat.HTML);
```

---

# Spire.Presentation Slide to SVG Conversion
## Convert a PowerPoint slide to SVG format
```java
//Create PPT document
Presentation presentation = new Presentation();

//Load PPT file from disk
presentation.loadFromFile("data/OneSlideToSVG.pptx");

//Convert the second slide to SVG
byte[] svgByte = presentation.getSlides().get(1).SaveToSVG();
```

---

# Spire.Presentation Default Font Setting
## Set and reset default font for presentations
```java
//Set the default font
Presentation.setDefaultFontName("Bell MT");
//Reset the default font
Presentation.resetDefaultFontName();
```

---

# Spire Presentation Specific Slide to PDF Conversion
## Convert a specific slide from PowerPoint to PDF format
```java
//Create PPT document
Presentation presentation = new Presentation();

//Load the PPT document from disk
presentation.loadFromFile("data/ChangeSlidePosition.pptx");

//Get the second slide
ISlide slide = presentation.getSlides().get(1);

//Save the second slide to PDF
slide.SaveToFile("output/specificSlideToPDF_result.pdf", FileFormat.PDF);
```

---

# Spire.Presentation Slide to SVG Conversion
## Convert specific PowerPoint slides to SVG format
```java
// Create a new Presentation object
Presentation ppt = new Presentation();

// Load the PowerPoint file
ppt.loadFromFile(inputFile);

// Save specified slides (from index 0 to 1) as SVG format
List<byte[]> bytes = ppt.saveToSVG(0, 1);

// Dispose of the Presentation object to release resources
ppt.dispose();
```

---

# Spire.Presentation Font Directory Specification
## Specify custom font directory for presentation
```java
Presentation ppt = new Presentation();
//Specify font directory
ppt.setCustomFontsFolder("data/Fonts/");
```

---

# Spire.Presentation Slide Conversion
## Convert specific slides from PowerPoint to PDF
```java
// Create a new Presentation object
Presentation ppt = new Presentation();

// Load the PowerPoint file
ppt.loadFromFile(inputFile);

// Save the specified slide (from index 1 to 1) to PDF file format
ppt.saveToFile(1, 1, outputFile, FileFormat.PDF);

// Dispose of the Presentation object to release resources
ppt.dispose();
```

---

# Spire Presentation to HTML Conversion
## Convert PowerPoint presentation to HTML format
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Load file
ppt.loadFromFile("data/Conversion.pptx");

//Save the document to HTML format
String result = "output/ToHTML.html";
ppt.saveToFile(result, FileFormat.HTML);
```

---

# Spire Presentation to Image Conversion
## Convert PowerPoint slides to PNG images
```java
public class toImage {
    public static void main(String[] args) throws Exception{
        Presentation ppt = new Presentation();
        ppt.loadFromFile(inputFile);
        //Save PPT document to images
        for (int i = 0; i < ppt.getSlides().getCount(); i++) {
            BufferedImage image = ppt.getSlides().get(i).saveAsImage();
            String fileName = outputFile + "/" + String.format("ToImage-%1$s.png", i);
            ImageIO.write(image, "PNG",new File(fileName));
        }
        ppt.dispose();
    }
}
```

---

# spire presentation to pdf conversion
## convert PowerPoint presentation to PDF format
```java
// Create a presentation object
Presentation ppt = new Presentation();

// Load the PowerPoint file
ppt.loadFromFile(inputFile);

// Save the presentation as PDF
ppt.saveToFile(outputFile, FileFormat.PDF);

// Dispose the presentation object
ppt.dispose();
```

---

# Spire Presentation PDF Conversion
## Convert PPT to PDF with specific page size
```java
Presentation ppt = new Presentation();

//Set A4 page size
ppt.getSlideSize().setType(SlideSizeType.A4);
```

---

# PPT to PPTX Conversion
## Convert PowerPoint presentation from PPT to PPTX format
```java
Presentation pt = new Presentation();
//Load the PPT file from disk
pt.loadFromFile(inputFile);
//Save the PPT document to PPTX file format
pt.saveToFile(outputFile, FileFormat.PPTX_2013);
pt.dispose();
```

---

# Spire Presentation Conversion
## Convert PPT slides to specific size images
```java
// Convert each slide to image with specific size
for (int i = 0; i < ppt.getSlides().getCount(); i++) {
    BufferedImage image = ppt.getSlides().get(i).saveAsImage(600,400);
}
```

---

# Spire.Presentation to SVG Conversion
## Convert PowerPoint presentation to SVG format
```java
String inputFile ="data/OneSlideToSVG.pptx";
String outputFile="output/";

Presentation ppt = new Presentation();
ppt.loadFromFile(inputFile);
ArrayList<byte[]> svgBytes =(ArrayList<byte[]>) ppt.saveToSVG();
for (int i = 0; i < svgBytes.size(); i++)
{
    byte[] bytes = svgBytes.get(i);
    FileOutputStream stream = new FileOutputStream(String.format(outputFile + "ToSVG-%d.svg", i));
    stream.write(bytes);
}
ppt.dispose();
```

---

# PowerPoint to SVGZ Conversion
## Convert PowerPoint presentation slides to SVGZ format
```java
// Create a new Presentation object
Presentation ppt = new Presentation();

// Load the PowerPoint file
ppt.loadFromFile("input_path");

// Save each slide as SVGZ format
List<byte[]> bytes = ppt.saveToSVGZ();

// Iterate through the saved SVGZ bytes and write them to individual files
for (int i = 0; i < bytes.size(); i++) {
    // Create a FileOutputStream for writing the SVGZ content to a file
    FileOutputStream fileOutputStream = new FileOutputStream("output_path" + "slide" + i + ".svgz");
    
    // Write the SVGZ content to the file
    fileOutputStream.write(bytes.get(i));
    
    // Flush and close the FileOutputStream
    fileOutputStream.flush();
    fileOutputStream.close();
}

// Dispose of the Presentation object to release resources
ppt.dispose();
```

---

# Spire Presentation Conversion
## Convert PowerPoint presentation to XPS format
```java
Presentation ppt = new Presentation();
ppt.loadFromFile(inputFile);

//Save the PPT to XPS file format
ppt.saveToFile(outputFile, FileFormat.XPS);
ppt.dispose();
```

---

# Spire.Presentation Image Stream
## Append image to presentation using stream
```java
// Use streaming to load image
FileInputStream fileInputStream = new FileInputStream("imagePath");

// Append a new image to replace an existing image
// Assuming ppt is an existing Presentation object
IImageData imageData = ppt.getImages().append(fileInputStream);
SlidePicture slidePicture = (SlidePicture) ppt.getSlides().get(0).getShapes().get(0);
slidePicture.getPictureFill().getPicture().setEmbedImage(imageData);
```

---

# Spire.Presentation Image Resizing
## Change the size of embedded images in a PowerPoint presentation
```java
float scale = 0.5f; // Scale factor for resizing images

// Iterate through all slides
for (int i = 0; i < ppt.getSlides().getCount(); i++) {
    ISlide slide = ppt.getSlides().get(i);
    
    // Iterate through all shapes in the slide
    for(int j = 0; j < slide.getShapes().getCount(); j++) {
        IShape shape = slide.getShapes().get(j);
        
        // Check if the shape is an embedded image
        if (shape instanceof IEmbedImage) {
            IEmbedImage image = (IEmbedImage)shape;
            
            // Resize the image
            image.setWidth(image.getWidth() * scale);
            image.setHeight(image.getHeight() * scale);
        }
    }
}
```

---

# Spire Presentation Image Cropping
## Crop image in presentation slide
```java
//Get first shape in first slide
IShape shape=presentation.getSlides().get(0).getShapes().get(0);
//If the shape is SlidePicture
if(shape instanceof SlidePicture)
{
    SlidePicture slidePicture= (SlidePicture) shape;
    //Crop the image
    slidePicture.crop(slidePicture.getLeft()+50f,slidePicture.getTop()+50f,100f,200f);
}
```

---

# Extract Images from Presentation
## Extract all images from a PowerPoint presentation
```java
// Extract image
BufferedImage image = ppt.getImages().get(i).getImage();
```

---

# Extract Images from PowerPoint Slide
## This code demonstrates how to extract images from specific shapes in a PowerPoint slide
```java
//Traverse all shapes in the second slide
for (int j = 0; j < ppt.getSlides().get(1).getShapes().getCount(); j++) {
    IShape shape = ppt.getSlides().get(1).getShapes().get(j);
    //It is the SlidePicture object
    if (shape instanceof SlidePicture) {
        SlidePicture ps = (SlidePicture) shape;
        BufferedImage image = ps.getPictureFill().getPicture().getEmbedImage().getImage();
    }
    //It is the PictureShape object
    if (shape instanceof PictureShape) {
        PictureShape ps = (PictureShape) shape;
        BufferedImage image = ps.getEmbedImage().getImage();
    }
}
```

---

# Get Clipping Information of Image
## Retrieve cropping and position information of an image in a presentation slide
```java
IShape shape = ppt.getSlides().get(0).getShapes().get(0);
if (shape instanceof SlidePicture) {
    SlidePicture picture = (SlidePicture)shape;
    // Get the cropped position
    Rectangle2D cropPosition = picture.getPictureFill().getCropPosition();
    // Get the position of picture
    Rectangle2D picPosition = picture.getPictureFill().getPicturePosition();
}
```

---

# Spire Presentation Image Path Extraction
## Get relative paths of images in a PowerPoint presentation
```java
//Create PPT document
Presentation ppt = new Presentation();
//Load document from disk
ppt.loadFromFile("data/RemoveImages.pptx");
//Get image collection
ImageCollection images = ppt.getImages();
for (int i = 0; i < images.size(); i++){
    IImageData imageData = images.get(i);
    //Get image relative path
    String path = imageData.getRelativePath();
}
```

---

# Spire.Presentation Image Insertion
## Core code for inserting an image into a PowerPoint presentation
```java
//Insert image to PPT
Rectangle2D.Double rect1 = new Rectangle2D.Double(presentation.getSlideSize().getSize().getWidth() / 2 - 280, 140, 120, 120);
IEmbedImage image = presentation.getSlides().get(0).getShapes().appendEmbedImage(ShapeType.RECTANGLE, imageFile, rect1);
image.getLine().setFillType(FillFormatType.NONE);
```

---

# Remove Images from Presentation
## Remove all images from a slide in a PowerPoint presentation
```java
//Get the first slide
ISlide slide = presentation.getSlides().get(0);

for (int i = slide.getShapes().getCount()-1; i >=0; i--)
{
    //Check if it is the SlidePicture object
    if (slide.getShapes().get(i) instanceof SlidePicture)
    {
        slide.getShapes().removeAt(i);
    }
}
```

---

# Image Frame Formatting in Presentation
## Set format properties for an image frame in a PowerPoint presentation
```java
//Create presentation and insert image
Presentation presentation = new Presentation();
Rectangle2D.Double rect1 = new Rectangle2D.Double(presentation.getSlideSize().getSize().getWidth() / 2 - 280, 140, 120, 120);
IEmbedImage image = presentation.getSlides().get(0).getShapes().appendEmbedImage(ShapeType.RECTANGLE, imageData, rect1);

//Set the formatting of the image frame
image.getLine().setFillType(FillFormatType.SOLID);
image.getLine().getSolidFillColor().setColor(Color.lightGray);
image.getLine().setWidth(5);
image.setRotation(45);
```

---

# spire.presentation image transparency
## Set transparency for an image in a presentation slide
```java
//Insert image to PPT
Rectangle2D.Double rect1 = new Rectangle2D.Double(200, 140, 120, 120);
IAutoShape shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, rect1);

shape.getLine().setFillType(FillFormatType.NONE);
shape.getFill().setFillType(FillFormatType.PICTURE);
shape.getFill().getPictureFill().setFillType(PictureFillType.STRETCH);
//Set transparency on image
shape.getFill().getPictureFill().getPicture().setTransparency(50);
```

---

# Spire.Presentation Image Update
## Update an existing image in a PowerPoint presentation
```java
//Get the first slide
ISlide slide = ppt.getSlides().get(0);

//Append a new image to replace an existing image
BufferedImage bufferedImage = (BufferedImage) ImageIO.read(new File("data/InsertImage.png"));
IImageData imageData = ppt.getImages().append(bufferedImage);

//Replace the image which title is "image1" with the new image
for (int j = 0; j < slide.getShapes().getCount(); j++) {
    IShape shape = slide.getShapes().get(j);
    if (shape instanceof SlidePicture) {
        if (shape.getAlternativeTitle().equals("image1")) {
            SlidePicture pic = (SlidePicture)shape;
            pic.getPictureFill().getPicture().setEmbedImage(imageData);
        }
    }
}
```

---

# Add Image in Table Cell
## Insert an image into a specific cell of a table in a PowerPoint presentation
```java
//Get the first shape
ITable table = (ITable) ((ppt.getSlides().get(0).getShapes().get(0) instanceof ITable) ? ppt.getSlides().get(0).getShapes().get(0) : null);

//Load the image and insert it into table cell
IImageData pptImg = ppt.getImages().append(image);
table.get(1, 1).getFillFormat().setFillType(FillFormatType.PICTURE);
table.get(1, 1).getFillFormat().getPictureFill().getPicture().setEmbedImage(pptImg);
table.get(1, 1).getFillFormat().getPictureFill().setFillType(PictureFillType.STRETCH);
```

---

# PowerPoint Table Row Addition
## Add a new row to an existing table in PowerPoint by cloning and appending
```java
//Get the table within the PowerPoint document.
ITable table = (ITable)presentation.getSlides().get(0).getShapes().get(0);

//Get the second row.
TableRow row = table.getTableRows().get(1);

//Clone the row and add it to the end of table.
table.getTableRows().append(row);
int rowCount = table.getTableRows().getCount();

//Get the last row.
TableRow lastRow = table.getTableRows().get(rowCount-1);

//Set new data of the first cell of last row.
lastRow.get(0).getTextFrame().setText("The first added cell");

//Set new data of the second cell of last row.
lastRow.get(1).getTextFrame().setText("The second added cell");
```

---

# Spire Presentation Table Column Adjustment
## Adjust column width based on text content
```java
//Get the table object from the first slide
ITable table = (ITable) ppt.getSlides().get(0).getShapes().get(0);

//Adjust the first column width of table by text width
table.getColumnsList().get(0).adjustColumnByTextWidth();
```

---

# Spire.Presentation Table Cloning
## Clone rows and columns in a presentation table
```java
// Clone row 1 at end of table
table.getTableRows().append(table.getTableRows().get(0));

// Clone row 2 as the 4th row of table
table.getTableRows().insert(3, table.getTableRows().get(1));

// Clone column 1 at end of table
table.getColumnsList().add(table.getColumnsList().get(0));

// Clone the 2nd column at 4th column index
table.getColumnsList().insert(3, table.getColumnsList().get(1));
```

---

# Spire.Presentation Table Cloning
## Clone a table from one slide to another
```java
//Get the table
ITable table = (ITable)ppt.getSlides().get(0).getShapes().get(0);

//Clone the table of the first ppt to the second ppt
slide.getShapes().appendTable(60,60,table);
```

---

# Spire.Presentation Table Creation
## Create and format a table in PowerPoint presentation
```java
//Create a PPT document
Presentation presentation = new Presentation();

//Define table dimensions
Double[] widths = new Double[]{100d, 100d, 150d, 100d, 100d};
Double[] heights = new Double[]{15d, 15d, 15d, 15d, 15d, 15d, 15d, 15d, 15d, 15d, 15d, 15d, 15d};

//Add new table to PPT
ITable table = presentation.getSlides().get(0).getShapes().appendTable((float) presentation.getSlideSize().getSize().getWidth() / 2 - 275, 90, widths, heights);

//Add data to table
for (int i = 0; i < 13; i++) {
    for (int j = 0; j < 5; j++) {
        //Set the Font
        table.get(j, i).getTextFrame().getParagraphs().get(0).getTextRanges().get(0).setLatinFont(new TextFont("Arial Narrow"));
    }
}

//Set the alignment of the first row to Center
for (int i = 0; i < 5; i++) {
    table.get(i, 0).getTextFrame().getParagraphs().get(0).setAlignment(TextAlignmentType.CENTER);
}

//Set the style of table
table.setStylePreset(TableStylePreset.LIGHT_STYLE_3_ACCENT_1);
```

---

# PowerPoint Table Row and Column Distribution
## Distribute rows and columns evenly in a PowerPoint table
```java
//Get the first slide
ISlide slide = ppt.getSlides().get(0);

//Get the first table
ITable table = (ITable) slide.getShapes().get(0);

//distribute rows
table.distributeRows(1,3);

//distribute columns
table.distributeColumns(0,3);
```

---

# Spire.Presentation Table Editing
## Edit table data and style in a PowerPoint presentation
```java
//Create a PPT document
Presentation presentation = new Presentation();

ITable table = null;

//Get the table in PowerPoint document
for (int i = 0; i < presentation.getSlides().get(0).getShapes().getCount(); i++)
{
    IShape shape = presentation.getSlides().get(0).getShapes().get(i);
    if (shape instanceof ITable)
    {
        table = (ITable)shape;
        //Change the style of table
        table.setStylePreset(TableStylePreset.LIGHT_STYLE_1_ACCENT_2);
        for (int j = 0; j < table.getColumnsList().getCount(); j++)
        {
            //Replace the data in cell
            table.get(j,2).getTextFrame().setText("New Data");
            
            //Set the highlight color
            table.get(j,2).getTextFrame().getTextRange().getHighlightColor().setColor(Color.lightGray);
        }
    }
}
```

---

# Fill Table Cells with Color
## This code demonstrates how to fill all cells in a PowerPoint table with a solid color.
```java
//Get the table in PowerPoint document
for (int i = 0; i < presentation.getSlides().get(0).getShapes().getCount(); i++)
{
    IShape shape = presentation.getSlides().get(0).getShapes().get(i);
    if (shape instanceof ITable)
    {
        ITable table = (ITable)shape;

        for (int j = 0; j < table.getTableRows().getCount(); j++)
        {
            TableRow row = table.getTableRows().get(j);
            for (int a = 0; a < row.getCount(); a++)
            {
                row.get(a).getFillFormat().setFillType(FillFormatType.SOLID);
                row.get(a).getFillFormat().getSolidColor().setColor(Color.pink);
            }
        }
    }
}
```

---

# Fill Table Row with Color
## This code demonstrates how to fill a particular row in a PowerPoint table with a specific color
```java
//Get the table in PowerPoint document.
for (int i = 0; i < presentation.getSlides().get(0).getShapes().getCount(); i++) {
    IShape shape = presentation.getSlides().get(0).getShapes().get(i);
    if (shape instanceof ITable) {
        table = (ITable) shape;

        //Fill particular table row with color.
        TableRow row = table.getTableRows().get(1);
        for (int a = 0; a < row.getCount(); a++) {
            row.get(a).getFillFormat().setFillType(FillFormatType.SOLID);
            row.get(a).getFillFormat().getSolidColor().setColor(Color.pink);
        }
    }
}
```

---

# Spire.Presentation Identify Merged Cells
## Identify merged cells in PowerPoint table
```java
//Create a PPT document
Presentation presentation = new Presentation();

//Load the file from disk.
presentation.loadFromFile("data/MergedCellInTable.pptx");

ITable table = null;

//Get the table in PowerPoint document.
for (int i = 0; i < presentation.getSlides().get(0).getShapes().getCount();i ++) {
    IShape shape = presentation.getSlides().get(0).getShapes().get(i);
    if (shape instanceof ITable) {
        table = (ITable) shape;

        //Identify merged cells in the table
        for (int j = 0; j < table.getTableRows().getCount(); j++) {
            TableRow row = table.getTableRows().get(j);
            for (int a = 0; a < row.getCount(); a++) {
                //Check if cell is merged
                if (row.get(a).getRowSpan() > 1 || row.get(a).getColSpan() > 1) {
                    //Display merged cell information
                    System.out.println("The cell " + j + ":" + a + "is a part of merged cell with RowSpan=" + row.get(a).getRowSpan() + " and ColSpan=" + row.get(a).getColSpan() + " starting from Cell " + row.get(a).getFirstRowIndex() + " : " + row.get(a).getFirstColumnIndex());
                }
            }
        }
    }
}
```

---

# Spire Presentation Table Aspect Ratio Locking
## Lock aspect ratio of tables in PowerPoint presentation
```java
//Get the table in PowerPoint document.
for (int i = 0; i < presentation.getSlides().get(0).getShapes().getCount();i ++) {
    IShape shape = presentation.getSlides().get(0).getShapes().get(i);
    if (shape instanceof ITable)
    {
        ITable table = (ITable) shape;
        //Lock aspect ratio
        table.getShapeLocking().setAspectRatioProtection(true);
    }
}
```

---

# Spire.Presentation Table Cell Merging
## Merge table cells in a PowerPoint presentation
```java
//Find the table in the first slide
ITable table = null;
for (Object shape : presentation.getSlides().get(0).getShapes()) {
    if (shape instanceof ITable) {
        table = (ITable) shape;
        //Merge the second row and third row of the first column
        table.mergeCells(table.get(0, 1), table.get(0, 2), false);
        table.mergeCells(table.get(3, 4), table.get(4, 4), true);
    }
}
```

---

# Spire.Presentation Table Manipulation
## Remove rows and columns from a PowerPoint table
```java
//Get the table in PPT document
ITable table = null;
for (Object shape : presentation.getSlides().get(0).getShapes()) {
    if (shape instanceof ITable) {
        table = (ITable) shape;
        //Remove the second column
        table.getColumnsList().removeAt(1, false);
        //Remove the second row
        table.getTableRows().removeAt(1, false);
    }
}
```

---

# Spire.Presentation Table Border Style Removal
## Remove border styles from all cells in a PowerPoint table
```java
//Get the table in PowerPoint document.
for (int i = 0; i < presentation.getSlides().get(0).getShapes().getCount();i ++)
{
    IShape shape = presentation.getSlides().get(0).getShapes().get(i);
    if (shape instanceof  ITable)
    {
        table = (ITable)shape;

        for (int j = 0; j < table.getTableRows().getCount(); j++)
        {
            TableRow row = table.getTableRows().get(j);
            for (int a = 0; a < row.getCount(); a ++)
            {
                row.get(a).getBorderTop().setFillType(FillFormatType.NONE);
                row.get(a).getBorderBottom().setFillType(FillFormatType.NONE);
                row.get(a).getBorderLeft().setFillType(FillFormatType.NONE);
                row.get(a).getBorderRight().setFillType(FillFormatType.NONE);
            }
        }
    }
}
```

---

# Remove Table from PowerPoint Slide
## This code demonstrates how to find and remove a table from a PowerPoint slide
```java
// Get the table in PowerPoint document
ITable table = null;
for (int i = 0; i < presentation.getSlides().get(0).getShapes().getCount(); i++) {
    IShape shape = presentation.getSlides().get(0).getShapes().get(i);
    if (shape instanceof ITable) {
        table = (ITable)shape;
        // Remove the table from the slide
        presentation.getSlides().get(0).getShapes().remove(table);
    }
}
```

---

# Spire Presentation Table Alignment
## Set horizontal, vertical, and combined alignment in table cells
```java
ITable table = null;
for (Object shape : presentation.getSlides().get(0).getShapes()) {
    if (shape instanceof ITable) {
        table = (ITable) shape;
        //Horizontal Alignment
        //Set the horizontal alignment for the cells in first column
        table.get(0, 1).getTextFrame().getParagraphs().get(0).setAlignment(TextAlignmentType.LEFT);
        table.get(0, 2).getTextFrame().getParagraphs().get(0).setAlignment(TextAlignmentType.CENTER);
        table.get(0, 3).getTextFrame().getParagraphs().get(0).setAlignment(TextAlignmentType.RIGHT);
        table.get(0, 4).getTextFrame().getParagraphs().get(0).setAlignment(TextAlignmentType.JUSTIFY);
        //Vertical Alignment
        //Set the vertical alignment for the cells in second column
        table.get(1, 1).setTextAnchorType(TextAnchorType.TOP);
        table.get(1, 2).setTextAnchorType(TextAnchorType.CENTER);
        table.get(1, 3).setTextAnchorType(TextAnchorType.BOTTOM);
        table.get(1, 4).setTextAnchorType(TextAnchorType.NONE);

        //Both orientations
        //Set the both horizontal and vertical alignment for the cells in the third column
        table.get(2, 1).getTextFrame().getParagraphs().get(0).setAlignment(TextAlignmentType.LEFT);
        table.get(2, 1).setTextAnchorType(TextAnchorType.TOP);
        table.get(2, 2).getTextFrame().getParagraphs().get(0).setAlignment(TextAlignmentType.RIGHT);
        table.get(2, 2).setTextAnchorType(TextAnchorType.CENTER);
        table.get(2, 3).getTextFrame().getParagraphs().get(0).setAlignment(TextAlignmentType.JUSTIFY);
        table.get(2, 3).setTextAnchorType(TextAnchorType.BOTTOM);
        table.get(2, 4).getTextFrame().getParagraphs().get(0).setAlignment(TextAlignmentType.CENTER);
        table.get(2, 4).setTextAnchorType(TextAnchorType.TOP);
    }
}
```

---

# spire presentation table borders
## set borders for existing table in PowerPoint
```java
//Get the table within the PowerPoint document.
ITable table = (ITable)presentation.getSlides().get(0).getShapes().get(0);

//Set the border type as Inside and the border color as blue.
table.setTableBorder(TableBorderType.Inside, 1, Color.blue);
```

---

# Spire.Presentation Table Borders
## Set borders for newly created tables in PowerPoint presentations

```java
Presentation presentation = new Presentation();

// Define table dimensions
Double[] tableWidth = new Double[] { 100d, 100d, 100d, 100d, 100d };
Double[] tableHeight = new Double[] { 20d, 20d, 20d };

for (TableBorderType e : TableBorderType.values()) {
    // Add a table to the presentation slide
    ITable itable = presentation.getSlides().append().getShapes().appendTable(100, 100, tableWidth, tableHeight);
    
    // Add text to table cells
    itable.getTableRows().get(0).get(0).getTextFrame().setText("Row");
    itable.getTableRows().get(1).get(0).getTextFrame().setText("Column");
    
    // Set the border type, border width and color for the table
    itable.setTableBorder(TableBorderType.valueOf(e.toString()), 1.5, Color.red);
}
```

---

# Spire Presentation Table Formatting
## Set row height and column width for a table in a PowerPoint presentation
```java
//Get the table
ITable table = null;
for (Object shape : ppt.getSlides().get(0).getShapes()) {
    if (shape instanceof ITable) {
        table = (ITable) shape;
        //Set the height for the rows
        table.getTableRows().get(0).setHeight(100);
        table.getTableRows().get(1).setHeight(80);
        table.getTableRows().get(2).setHeight(60);
        table.getTableRows().get(3).setHeight(40);
        table.getTableRows().get(4).setHeight(20);
        //Set the column width
        table.getColumnsList().get(0).setWidth(60);
        table.getColumnsList().get(1).setWidth(80);
        table.getColumnsList().get(2).setWidth(120);
        table.getColumnsList().get(3).setWidth(140);
        table.getColumnsList().get(4).setWidth(160);
    }
}
```

---

# Spire.Presentation Table Border Style
## Set border styles for tables in PowerPoint presentation
```java
// Find the table by looping through all the slides, and then set borders for it.
for (int b = 0; b < presentation.getSlides().getCount(); b++) {
    ISlide slide = presentation.getSlides().get(b);
    for (int i = 0; i < slide.getShapes().getCount(); i++) {
        IShape shape = slide.getShapes().get(i);
        if (shape instanceof ITable) {
            table = (ITable) shape;

            for (int j = 0; j < table.getTableRows().getCount(); j++) {
                TableRow row = table.getTableRows().get(j);
                for (int a = 0; a < row.getCount(); a++) {
                    Cell cell = row.get(a);
                    cell.getBorderTop().setFillType(FillFormatType.SOLID);
                    cell.getBorderBottom().setFillType(FillFormatType.SOLID);
                    cell.getBorderLeft().setFillType(FillFormatType.SOLID);
                    cell.getBorderRight().setFillType(FillFormatType.SOLID);
                }
            }
        }
    }
}
```

---

# Spire Presentation Table Styling
## Set table style preset in PowerPoint presentation
```java
ITable table = null;

//Get the table in PowerPoint document.
for (int i = 0; i < presentation.getSlides().get(0).getShapes().getCount();i ++)
{
    IShape shape = presentation.getSlides().get(0).getShapes().get(i);
    if (shape instanceof  ITable)
    {
        table = (ITable)shape;
        //Set the style of table.
        table.setStylePreset(TableStylePreset.MEDIUM_STYLE_1_ACCENT_2);
    }
}
```

---

# Spire.Presentation Table Text Formatting
## Set text formatting for table cells in PowerPoint presentation
```java
//Get the table in PowerPoint document
for (int i = 0; i < presentation.getSlides().get(0).getShapes().getCount(); i++) {
    IShape shape = presentation.getSlides().get(0).getShapes().get(i);
    if (shape instanceof ITable) {
        ITable table = (ITable)shape;
        Cell cell1 = table.getTableRows().get(0).get(0);
        //Set table cell's text alignment type 
        cell1.setTextAnchorType(TextAnchorType.TOP);
        //Set italic style
        cell1.getTextFrame().getTextRange().getFormat().isItalic(TriState.TRUE);

        Cell cell2 = table.getTableRows().get(1).get(0);
        //Set table cell's foreground color
        cell2.getTextFrame().getTextRange().getFill().setFillType(FillFormatType.SOLID);
        cell2.getTextFrame().getTextRange().getFill().getSolidColor().setColor(Color.green);
        //Set table cell's background color
        cell2.getFillFormat().setFillType(FillFormatType.SOLID);
        cell2.getFillFormat().getSolidColor().setColor(Color.lightGray);

        Cell cell3 = table.getTableRows().get(2).get(2);
        //Set table cell's font and font size
        cell3.getTextFrame().getTextRange().setFontHeight(12);
        cell3.getTextFrame().getTextRange().setLatinFont(new TextFont("Arial Black"));
        cell3.getTextFrame().getTextRange().getHighlightColor().setColor(Color.yellow);

        Cell cell4 = table.getTableRows().get(2).get(1);
        //Set table cell's margin and borders
        cell4.setMarginLeft(20);
        cell4.setMarginTop(30);
        cell4.getBorderTop().setFillType(FillFormatType.SOLID);
        cell4.getBorderTop().getSolidFillColor().setColor(Color.red);
        cell4.getBorderBottom().setFillType(FillFormatType.SOLID);
        cell4.getBorderBottom().getSolidFillColor().setColor(Color.red);
        cell4.getBorderLeft().setFillType(FillFormatType.SOLID);
        cell4.getBorderLeft().getSolidFillColor().setColor(Color.red);
        cell4.getBorderRight().setFillType(FillFormatType.SOLID);
        cell4.getBorderRight().getSolidFillColor().setColor(Color.red);
    }
}
```

---

# Spire.Presentation Table Cell Splitting
## Split a specific table cell in a PowerPoint presentation
```java
//Create a PowerPoint document.
Presentation presentation = new Presentation();

//Get the table within the PowerPoint document.
ITable table = (ITable)presentation.getSlides().get(0).getShapes().get(0);

//Split cell [1, 2] into 3 rows and 2 columns.
table.getTableRows().get(1).get(2).split(3,2);
```

---

# PowerPoint Table Cell Traversal
## Traverse through cells in a PowerPoint table
```java
//Get the table in PowerPoint document.
for (int i = 0; i < presentation.getSlides().get(0).getShapes().getCount(); i++) {
    IShape shape = presentation.getSlides().get(0).getShapes().get(i);
    if (shape instanceof ITable) {
        ITable table = (ITable) shape;

        // Traverse through rows of the table
        for (int j = 0; j < table.getTableRows().getCount(); j++) {
            TableRow row = table.getTableRows().get(j);
            
            // Traverse through cells of the row
            for (int a = 0; a < row.getCount(); a++) {
                Cell cell = row.get(a);
                // Get text from cell
                cell.getTextFrame().getText();
            }
        }
    }
}
```

---

# Spire.Presentation Hyperlink to Image
## Add hyperlink to an image in a PowerPoint presentation
```java
//Create a PPT document
Presentation presentation = new Presentation();

//Get the first slide
ISlide slide = presentation.getSlides().get(0);

//Add image to slide
Rectangle2D.Double rect = new Rectangle2D.Double(480, 350, 160, 160);
IEmbedImage image = slide.getShapes().appendEmbedImage(ShapeType.RECTANGLE, imageFile, rect);

//Add hyperlink to the image
ClickHyperlink hyperlink = new ClickHyperlink("https://www.e-iceblue.com");
image.setClick(hyperlink);
```

---

# Add Hyperlink to Text in Presentation
## Demonstrates how to add hyperlink to text in a PowerPoint presentation
```java
IAutoShape shape = (IAutoShape)presentation.getSlides().get(0).getShapes().get(0);
ParagraphEx tp = shape.getTextFrame().getParagraphs().get(0);
String temp = tp.getText();

//Clear all text.
tp.getTextRanges().clear();

//Split the original text.
String[] strSplit = temp.split("Spire.Presentation");

//Add new text.
PortionEx tr = new PortionEx(strSplit[0]);
tp.getTextRanges().append(tr);

//Add the hyperlink.
tr = new PortionEx("Spire.Presentation");
tr.getClickAction().setAddress("http://www.e-iceblue.com/Introduce/presentation-for-net-introduce.html");
tp.getTextRanges().append(tr);
```

---

# Change Hyperlink Color in PowerPoint
## Change the color of hyperlinks in a PowerPoint presentation using Spire.Presentation
```java
//Create a PowerPoint document
Presentation presentation = new Presentation();

//Get the first slide
ISlide slide = presentation.getSlides().get(0);

//Get the theme of the slide
Theme theme = slide.getTheme();

//Change the color of hyperlink to red
theme.getColorScheme().getHyperlinkColor().setColor(Color.red);
```

---

# spire presentation hyperlink outline style
## create hyperlink with custom outline style in PowerPoint presentation
```java
//Create PPT document
Presentation presentation = new Presentation();

//Add new shape to PPT document
Rectangle rec = new Rectangle((int)presentation.getSlideSize().getSize().getWidth() / 2 - 250, 120, 400, 100);
IAutoShape shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, rec);

shape.getFill().setFillType(FillFormatType.NONE);
shape.getLine().setFillType(FillFormatType.NONE);

//Add some paragraphs with hyperlinks
ParagraphEx para1 = new ParagraphEx();
PortionEx tr1 = new PortionEx();
tr1.setText("Click to know more about Spire.Presentation.");
tr1.getClickAction().setAddress("https://www.e-iceblue.com/Introduce/presentation-for-java.html");
para1.getTextRanges().append(tr1);

tr1.getFormat().isItalic(TriState.TRUE);
tr1.getFormat().setFontMinSize(20);

//Set the outline format of text range
tr1.getTextLineFormat().getFillFormat().setFillType(FillFormatType.SOLID);
tr1.getTextLineFormat().getFillFormat().getSolidFillColor().setColor(Color.lightGray);
tr1.getTextLineFormat().setJoinStyle(LineJoinType.ROUND);
tr1.getTextLineFormat().setWidth(2);

//Add the paragraph to shape
shape.getTextFrame().getParagraphs().append(para1);
```

---

# Spire.Presentation Hyperlinks
## Creating hyperlinks in a PowerPoint presentation
```java
//Add new shape to PPT document
Rectangle2D.Double rec = new Rectangle2D.Double(presentation.getSlideSize().getSize().getWidth() / 2 - 255, 120, 500, 280);
IAutoShape shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, rec);
shape.getFill().setFillType(FillFormatType.NONE);
shape.getLine().setWidth(0);

//Add some paragraphs with hyperlinks
ParagraphEx para1 = new ParagraphEx();
PortionEx tr1 = new PortionEx();
tr1.setText("E-iceblue");
para1.getTextRanges().append(tr1);
shape.getTextFrame().getParagraphs().append(para1);
para1.setAlignment(TextAlignmentType.CENTER);
//Set the font and fill style of text
tr1.getFill().setFillType(FillFormatType.SOLID);
tr1.getFill().getSolidColor().setColor(Color.blue);
shape.getTextFrame().getParagraphs().append(new ParagraphEx());

//Add some paragraphs with hyperlinks
ParagraphEx para2 = new ParagraphEx();
PortionEx tr2 = new PortionEx();
tr2.setText("Click to know more about Spire.Presentation.");
tr2.getClickAction().setAddress("https://www.e-iceblue.com/Introduce/presentation-for-java.html");
para2.getTextRanges().append(tr2);
shape.getTextFrame().getParagraphs().append(para2);
shape.getTextFrame().getParagraphs().append(new ParagraphEx());

ParagraphEx para3 = new ParagraphEx();
PortionEx tr3 = new PortionEx();
tr3.setText("Click to visit E-iceblue Home page.");
tr3.getClickAction().setAddress("https://www.e-iceblue.com/");
para3.getTextRanges().append(tr3);
shape.getTextFrame().getParagraphs().append(para3);
shape.getTextFrame().getParagraphs().append(new ParagraphEx());

ParagraphEx para4 = new ParagraphEx();
PortionEx tr4 = new PortionEx();
tr4.setText("Click to go to the forum to raise questions.");
tr4.getClickAction().setAddress("https://www.e-iceblue.com/forum/components-f5.html");
para4.getTextRanges().append(tr4);
shape.getTextFrame().getParagraphs().append(para4);
shape.getTextFrame().getParagraphs().append(new ParagraphEx());

ParagraphEx para5 = new ParagraphEx();
PortionEx tr5 = new PortionEx();
tr5.setText("Click to contact our sales team via email. ");
tr5.getClickAction().setAddress("mailto:sales@e-iceblue.com");
para5.getTextRanges().append(tr5);
shape.getTextFrame().getParagraphs().append(para5);
shape.getTextFrame().getParagraphs().append(new ParagraphEx());

ParagraphEx para6 = new ParagraphEx();
PortionEx tr6 = new PortionEx();
tr6.setText("Click to contact our support team via email. ");
tr6.getClickAction().setAddress("mailto:support@e-iceblue.com");
para6.getTextRanges().append(tr6);
shape.getTextFrame().getParagraphs().append(para6);

for (Object para : shape.getTextFrame().getParagraphs()) {
    ParagraphEx paragraph = (ParagraphEx) para;
    String text = paragraph.getText();
    if (text != null && text.length() != 0) {
        paragraph.getTextRanges().get(0).setLatinFont(new TextFont("Lucida Sans Unicode"));
        paragraph.getTextRanges().get(0).setFontHeight(20);
    }
}
```

---

# Spire.Presentation hyperlink to specific slide
## Create a hyperlink that links to a specific slide in a PowerPoint presentation
```java
//Create PPT document
Presentation presentation = new Presentation();

//Append a slide to it.
presentation.getSlides().append();

//Add new shape to PPT document
Rectangle rec = new Rectangle((int)presentation.getSlideSize().getSize().getWidth() / 2 - 250, 120, 400, 100);
IAutoShape shape = presentation.getSlides().get(1).getShapes().appendShape(ShapeType.RECTANGLE, rec);

shape.getFill().setFillType(FillFormatType.NONE);
shape.getLine().setFillType(FillFormatType.NONE);
shape.getTextFrame().setText("Jump to the first slide");

//Create a hyperlink based on the shape and the text on it, linking to the first slide.
ClickHyperlink hyperlink = new ClickHyperlink(presentation.getSlides().get(0));
shape.setClick(hyperlink);
shape.getTextFrame().getTextRange().setClickAction(hyperlink);
```

---

# Spire Presentation Hyperlink
## Create hyperlink to last viewed slide
```java
//Create a shape
IAutoShape autoShape = ppt.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, new Rectangle2D.Double(100, 100, 100, 100));
//Link to recently viewed slide show
ClickHyperlink lastViewedSlide = ClickHyperlink.getLastViewedSlide();
autoShape.setClick(lastViewedSlide);
```

---

# Spire.Presentation Hyperlink Modification
## Modify hyperlink address and text in a PowerPoint presentation
```java
// Create a PPT document
Presentation presentation = new Presentation();

// Get the shape we want to edit it.
IAutoShape shape = (IAutoShape)presentation.getSlides().get(0).getShapes().get(0);

// Edit the link text and the target URL.
shape.getTextFrame().getTextRange().getClickAction().setAddress("http://www.e-iceblue.com");
shape.getTextFrame().getTextRange().setText("E-iceblue");
```

---

# Spire.Presentation Hyperlink Removal
## Remove hyperlink from a shape in a presentation
```java
//Get the shape and its text with hyperlink.
IAutoShape shape = (IAutoShape)presentation.getSlides().get(0).getShapes().get(0);

//Set null to remove the hyperlink.
shape.getTextFrame().getTextRange().setClickAction(null);
```

---

# Extract Video from PowerPoint
## Extract embedded videos from PowerPoint presentation slides
```java
//Create PPT document
Presentation presentation = new Presentation();

//Load the PPT document from disk.
presentation.loadFromFile("data/video.pptx");

//Counter for video files
int i = 0;

//Traverse all the slides of PPT file
for (Object slideObj : presentation.getSlides())
{
    ISlide slide=(ISlide)slideObj;
    //Traverse all the shapes of slides
    for (Object shapeObj : slide.getShapes())
    {
        IShape shape=(IShape)shapeObj;
        //If shape is IVideo
        if (shape instanceof IVideo)
        {
            //Save the video
            String result = "output/Video" + i + ".avi";
            ((IVideo)shape).getEmbeddedVideoData().saveToFile(result);
            i++;
        }
    }
}
```

---

# spire presentation audio video part name
## Get audio and video part names from PowerPoint presentation
```java
// Loop through all slides
for (int i = 0; i < ppt.getSlides().getCount(); i++) {
    // Loop through all shapes
    for (int j = 0; j < ppt.getSlides().get(i).getShapes().getCount(); j++) {
        // Get specified shape
        IShape shape = ppt.getSlides().get(i).getShapes().get(j);
        // If shape is IAudio
        if (shape instanceof IAudio) {
            // Get IAudio name
            String name = ((IAudio) shape).getData().getPartName();
            // If shape is IVideo
        } else if (shape instanceof IVideo) {
            // Get IVideo name
            String name = ((IVideo) shape).getEmbeddedVideoData().getPartName();
        }
    }
}
```

---

# Spire.Presentation Audio Insertion
## Insert audio into a PowerPoint presentation
```java
//Add title
Rectangle2D.Double rec_title = new Rectangle2D.Double(50, 240, 160, 50);
IAutoShape shape_title = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, rec_title);
shape_title.getLine().setFillType(FillFormatType.NONE);

shape_title.getFill().setFillType(FillFormatType.NONE);
ParagraphEx para_title = new ParagraphEx();
para_title.setText("Audio:");
para_title.setAlignment(TextAlignmentType.CENTER);
para_title.getTextRanges().get(0).setLatinFont(new TextFont("Myriad Pro Light"));
para_title.getTextRanges().get(0).setFontHeight(32);
para_title.getTextRanges().get(0).isBold(TriState.TRUE);
para_title.getTextRanges().get(0).getFill().setFillType(FillFormatType.SOLID);
para_title.getTextRanges().get(0).getFill().getSolidColor().setColor(Color.gray);
shape_title.getTextFrame().getParagraphs().append(para_title);

//Insert audio into the document
Rectangle2D.Double audioRect = new Rectangle2D.Double(220, 240, 80, 80);
presentation.getSlides().get(0).getShapes().appendAudioMedia(audioFilePath, audioRect);
```

---

# Insert Video into PowerPoint Presentation
## Core functionality for inserting a video file into a PowerPoint slide
```java
// Create or load presentation
Presentation presentation = new Presentation();
presentation.loadFromFile(inputFile);

// Insert video into the document
Rectangle2D.Double videoRect = new Rectangle2D.Double(presentation.getSlideSize().getSize().getWidth() / 2 - 125, 240, 150, 150);
IVideo video = presentation.getSlides().get(0).getShapes().appendVideoMedia((new java.io.File(inputFile_1)).getAbsolutePath(), videoRect);
BufferedImage image = ImageIO.read(new File(imageFile));
video.getPictureFill().getPicture().setEmbedImage(presentation.getImages().append(image));
```

---

# Spire Presentation Sound Effect Extraction
## Extract sound effect properties from presentation slides
```java
//Create an instance of presentation document
Presentation ppt = new Presentation();

//Get the first slide
ISlide slide = ppt.getSlides().get(0);

//Get the audio in a time node TimeNodeAudio
TimeNodeAudioEx audio = slide.getTimeline().getMainSequence().get(0).getTimeNodeAudios()[0];

//Get the properties of the audio, such as sound name, volume or detect if it's mute
String soundName = audio.getSoundName();
int volume = audio.getVolume();
boolean isMute = audio.isMute();
```

---

# Spire.Presentation Video Play Mode
## Set play mode for videos in PowerPoint presentation
```java
//Find the video by looping through all the slides and set its play mode as auto.
for (Object slideObj : presentation.getSlides())
{
    ISlide slide=(ISlide)slideObj;
    for (Object shapeObj : slide.getShapes())
    {
        IShape shape=(IShape)shapeObj;
        if (shape instanceof IVideo)
        {
            ((IVideo)shape).setPlayMode(VideoPlayMode.AUTO);
        }
    }
}
```

---

# Spire.Presentation Audio Update
## Update audio in PowerPoint presentation
```java
//Get Audio collection
WavAudioCollection audios = ppt.getWavAudios();
//update Audio
IAudioData audioData = audios.append(data);

//Get the specified shape
IShape shape = ppt.getSlides().get(0).getShapes().get(3);
//If shape is IAudio
if(shape instanceof IAudio){
    //update Audio
    ((IAudio)shape).setData(audioData);
}
```

---

# Spire.Presentation Video Update
## Update video in PowerPoint presentation
```java
//Create PPT document
Presentation ppt = new Presentation();

//Get video collection
VideoCollection videos = ppt.getVideos();
VideoData videoData = videos.append(data);

//Get the specified shape
ISlide iSlide = ppt.getSlides().get(0);

//Traverse all the shapes of slides
for (Object shape : iSlide.getShapes()) {
    //If shape is IVideo
    if (shape instanceof IVideo) {
        IVideo video = (IVideo) shape;
        //Update video
        video.setEmbeddedVideoData(videoData);
    }
}
```

---

# Spire.Presentation Speaker Notes Management
## Add and retrieve speaker notes in PowerPoint slides
```java
// Create a PowerPoint document
Presentation presentation = new Presentation();

// Get the first slide in the PowerPoint document
ISlide slide = presentation.getSlides().get(0);

// Get the NotesSlide in the first slide, if there is no notes, we need to add it firstly
NotesSlide ns = slide.getNotesSlide();
if (ns == null)
{
    ns = slide.addNotesSlide();
}

// Add the text string as the notes
ns.getNotesTextFrame().setText("Speak notes added by Spire.Presentation");

// Get the speaker notes
String notesText = ns.getNotesTextFrame().getText();
```

---

# Spire Presentation Comment
## Add comment to presentation slide
```java
//Comment author
ICommentAuthor author = ppt.getCommentAuthors().addAuthor("E-iceblue", "comment:");

//Add comment
ppt.getSlides().get(0).addComment(author, "Add comment", new Point2D.Float(44, 19), new java.util.Date());
```

---

# Spire.Presentation Add Notes
## Add notes to a presentation slide with numbered bullet points
```java
//Add note slide
NotesSlide notesSlide = slide.addNotesSlide();

//Add paragraph in the notesSlide
ParagraphEx paragraph = new ParagraphEx();
paragraph.setText("Tips for making effective presentations:");
notesSlide.getNotesTextFrame().getParagraphs().append(paragraph);

paragraph = new ParagraphEx();
paragraph.setText("Use the slide master feature to create a consistent and simple design template.");
notesSlide.getNotesTextFrame().getParagraphs().append(paragraph);

//Set the bullet type for the paragraph in notesSlide
notesSlide.getNotesTextFrame().getParagraphs().get(1).setBulletType(TextBulletType.NUMBERED);
notesSlide.getNotesTextFrame().getParagraphs().get(1).setBulletStyle(NumberedBulletStyle.BULLET_ARABIC_PERIOD);

paragraph = new ParagraphEx();
paragraph.setText("Simplify and limit the number of words on each screen.");
notesSlide.getNotesTextFrame().getParagraphs().append(paragraph);
notesSlide.getNotesTextFrame().getParagraphs().get(2).setBulletType(TextBulletType.NUMBERED);
notesSlide.getNotesTextFrame().getParagraphs().get(2).setBulletStyle(NumberedBulletStyle.BULLET_ARABIC_PERIOD);

paragraph = new ParagraphEx();
paragraph.setText("Use contrasting colors for text and background.");
notesSlide.getNotesTextFrame().getParagraphs().append(paragraph);
notesSlide.getNotesTextFrame().getParagraphs().get(3).setBulletType(TextBulletType.NUMBERED);
notesSlide.getNotesTextFrame().getParagraphs().get(3).setBulletStyle(NumberedBulletStyle.BULLET_ARABIC_PERIOD);
```

---

# Spire Presentation Comment Management
## Delete and replace comments in a presentation
```java
//Replace the text in the comment
ppt.getSlides().get(0).getComments()[1].setText("Replace comment");

//Delete the third comment
ppt.getSlides().get(0).deleteComment(ppt.getSlides().get(0).getComments()[1]);
```

---

# Extract Comments from PowerPoint Presentation
## Extract comments from a slide in a PowerPoint presentation
```java
//Get all comments from the first slide.
Comment[] comments = presentation.getSlides().get(0).getComments();
```

---

# Spire.Presentation Slide Comments
## Extract comments from PowerPoint slides
```java
//Create a PPT document
Presentation presentation = new Presentation();

//Loop through comments
for (Object commentAuthorObj : presentation.getCommentAuthors())
{
    ICommentAuthor commentAuthor=(ICommentAuthor)commentAuthorObj;
    for (Object commentObj : commentAuthor.getCommentsList())
    {
        Comment comment=(Comment)commentObj;
        //Get comment information
        String commentText = comment.getText();
        String authorName = comment.getAuthorName();
        Date time = comment.getDateTime();
    }
}
```

---

# PowerPoint to SVG Conversion with Notes
## Convert PowerPoint slides to SVG format while retaining notes
```java
//Create a PowerPoint document.
Presentation presentation = new Presentation();

//Load the file from disk.
presentation.loadFromFile("data/Template_Ppt_5.pptx");

//Retain the notes while converting PowerPoint file to svg file.
presentation.setNoteRetained(true);

//Convert presentation slides to svg file.
List<byte[]> bytes = presentation.saveToSVG();
```

---

# Remove Note from Specific Slide
## This code demonstrates how to remove notes from a specific slide in a PowerPoint presentation
```java
//Get the first slide
ISlide slide = presentation.getSlides().get(0);

//Get note slide
NotesSlide note = slide.getNotesSlide();
//Clear note text
note.getNotesTextFrame().setText("");
```

---

# Remove Speaker Notes from PowerPoint Slide
## Remove speaker notes from a specific slide in a PowerPoint presentation
```java
//Get the first slide from the sample document.
ISlide slide = presentation.getSlides().get(0);

//Remove the first speak note.
slide.getNotesSlide().getNotesTextFrame().getParagraphs().removeAt(1);
```

---

# Presentation Header and Footer Configuration
## Code to configure header and footer settings in a PowerPoint presentation
```java
//Add footer
presentation.setFooterText("Demo of Spire.Presentation");

//Set the footer visible
presentation.setFooterVisible(true);

//Set the page number visible
presentation.setSlideNumberVisible(true);

//Set the date visible
presentation.setDateTimeVisible(true);
```

---

# Spire Presentation SmartArt Node Access
## Access and retrieve SmartArt child nodes and their properties
```java
for (Object shapeObj : presentation.getSlides().get(0).getShapes())
{
    IShape shape=(IShape)shapeObj;
    if (shape instanceof ISmartArt)
    {
        //Get the SmartArt and collect nodes
        ISmartArt sa = (ISmartArt)shape;
        ISmartArtNodeCollection nodes = sa.getNodes();

        int position = 0;
        //Access the parent node at position 0
        ISmartArtNode node = nodes.get(position);
        ISmartArtNode childnode;

        //Traverse through all child nodes inside SmartArt
        for (int i = 0; i < node.getChildNodes().getCount(); i++)
        {
            //Access SmartArt child node at index i
            childnode = node.getChildNodes().get(i);
            
            //Get SmartArt child node parameters
            String nodeText = childnode.getTextFrame().getText();
            int nodeLevel = childnode.getLevel();
            int nodePosition = childnode.getPosition();
        }
    }
}
```

---

# Spire Presentation SmartArt Access
## Extract SmartArt node properties (text, level, position)
```java
ISmartArtNode node;
for (Object shapeObj : presentation.getSlides().get(0).getShapes())
{
    IShape shape=(IShape)shapeObj;
    if (shape instanceof ISmartArt)
    {
        //Get the SmartArt
        ISmartArt sa = (ISmartArt)shape ;
        ISmartArtNodeCollection nodes = sa.getNodes();

        //Traverse through all nodes inside SmartArt
        for (int i = 0; i < nodes.getCount(); i++)
        {
            //Access SmartArt node at index i
            node = nodes.get(i);

            //Get the SmartArt node parameters
            String nodeText = node.getTextFrame().getText();
            int nodeLevel = node.getLevel();
            int nodePosition = node.getPosition();
        }
    }
}
```

---

# Access SmartArt Layout
## Access and retrieve SmartArt layout type from PowerPoint presentation
```java
//Iterate through shapes in the first slide
for (Object shapeObj : presentation.getSlides().get(0).getShapes())
{
    IShape shape=(IShape)shapeObj;
    if (shape instanceof ISmartArt)
    {
        //Get the SmartArt
        ISmartArt sa = (ISmartArt)shape;
        
        //Check SmartArt Layout
        String layout = sa.getLayoutType().toString();
    }
}
```

---

# Access SmartArt Child Node
## Retrieve and access specific child nodes in SmartArt shapes
```java
//Get SmartArt node collection
ISmartArtNodeCollection nodes = smartArt.getNodes();

//Access SmartArt node at index 0
ISmartArtNode node = nodes.get(0);

//Access SmartArt child node at index 1
ISmartArtNode childNode = node.getChildNodes().get(1);

//Get the SmartArt child node parameters
String nodeText = childNode.getTextFrame().getText();
int nodeLevel = childNode.getLevel();
int nodePosition = childNode.getPosition();
```

---

# SmartArt Node Management
## Add nodes to SmartArt by specific position
```java
for (Object shapeObj : presentation.getSlides().get(0).getShapes()){
    IShape shape=(IShape)shapeObj;
    if (shape instanceof ISmartArt){
        //Get the SmartArt and collect nodes
        ISmartArt smartArt = (ISmartArt)shape;
        int position = 0;
        //Add a new node at specific position
        ISmartArtNode node = smartArt.getNodes().addNodeByPosition(position);
        //Add text and set the text style
        node.getTextFrame().setText("New Node");
        node.getTextFrame().getTextRange().getFill().setFillType(FillFormatType.SOLID);
        node.getTextFrame().getTextRange().getFill().getSolidColor().setKnownColor(KnownColors.RED);

        //Get a node
        node  =  smartArt.getNodes().get(1);
        position = 1;
        //Add a new child node at specific position
        ISmartArtNode childNode = node.getChildNodes().addNodeByPosition(position);
        //Add text and set the text style
        childNode.getTextFrame().setText ("New child node");
        childNode.getTextFrame().getTextRange().getFill().setFillType(FillFormatType.SOLID);
        childNode.getTextFrame().getTextRange().getFill().getSolidColor().setKnownColor(KnownColors.BLUE);
    }
}
```

---

# Spire.Presentation Organization Chart
## Add organization charts to presentation slides
```java
//Create a PPT document
Presentation presentation = new Presentation();

//Get the first slide and insert Picture Organization Chart
ISlide slide0 =  presentation.getSlides().get(0);
slide0.getShapes().appendSmartArt(50, 50, 250, 250, SmartArtLayoutType.PICTURE_ORGANIZATION_CHART);

//Append a new slide and insert Name and Title Organization Chart
ISlide newSlide = presentation.getSlides().append();
newSlide.getShapes().appendSmartArt(50, 50, 250, 250, SmartArtLayoutType.NAME_AND_TITLE_ORGANIZATION_CHART);
```

---

# Spire.Presentation SmartArt Node
## Add a SmartArt node with text styling
```java
//Get the SmartArt
ISmartArt sa = (ISmartArt) ((ppt.getSlides().get(0).getShapes().get(0) instanceof ISmartArt) ? ppt.getSlides().get(0).getShapes().get(0) : null);

//Add a node
ISmartArtNode node = sa.getNodes().addNode();

//Add text and set the text style
node.getTextFrame().setText("AddText");
node.getTextFrame().getTextRange().getFill().setFillType(FillFormatType.SOLID);
node.getTextFrame().getTextRange().getFill().getSolidColor().setKnownColor(KnownColors.HOT_PINK);
```

---

# SmartArt Assistant Node Management
## Set SmartArt nodes as assistant nodes in a PowerPoint presentation
```java
// Get the SmartArt and collect nodes
ISmartArt smartArt = (ISmartArt)shape;
ISmartArtNodeCollection nodes = smartArt.getNodes();

// Traverse through all nodes inside SmartArt
for (int i = 0; i < nodes.getCount(); i++)
{
    // Access SmartArt node at index i
    ISmartArtNode node = nodes.get(i);
    // Check if node is assistant node
    if (!node.isAssistant())
    {
        // Set node as assistant node
        node.isAssistant(true);
    }
}
```

---

# SmartArt Node Text Modification
## Change text of a SmartArt node in PowerPoint presentation
```java
for (Object shapeObj: presentation.getSlides().get(0).getShapes())
{
    IShape shape=(IShape)shapeObj;
    if (shape instanceof ISmartArt)
    {
        //Get the SmartArt and collect nodes
        ISmartArt smartArt = (ISmartArt)shape;
        //Obtain the reference of a node by using its Index
        // select second root node
        ISmartArtNode node = smartArt.getNodes().get(1);
        // Set the text of the TextFrame
        node.getTextFrame().setText("Second root node");
    }
}
```

---

# Spire Presentation SmartArt Color Style
## Change SmartArt color style in PowerPoint presentation
```java
for (Object shapeObj: presentation.getSlides().get(0).getShapes())
{
    IShape shape=(IShape)shapeObj;
    if (shape instanceof ISmartArt)
    {
        //Get the SmartArt
        ISmartArt smartArt = (ISmartArt)shape;
        // Check SmartArt color type
        if (smartArt.getColorStyle().equals(SmartArtColorType.COLORED_FILL_ACCENT_1))
        {
            // Change SmartArt color type
            smartArt.setColorStyle(SmartArtColorType.COLORFUL_ACCENT_COLORS);
        }
    }
}
```

---

# Spire.Presentation SmartArt Style Change
## Change SmartArt shape style in PowerPoint presentation
```java
// Iterate through shapes in the first slide
for (Object shapeObj : presentation.getSlides().get(0).getShapes())
{
    IShape shape=(IShape)shapeObj;
    if (shape instanceof ISmartArt)
    {
        //Get the SmartArt
        ISmartArt smartArt = (ISmartArt)shape;
        //Check SmartArt style
        if (smartArt.getStyle() == SmartArtStyleType.SIMPLE_FILL)
        {
            //Change SmartArt Style
            smartArt.setStyle(SmartArtStyleType.CARTOON);
        }
    }
}
```

---

# Spire.Presentation SmartArt Shape Creation
## Create and customize SmartArt shapes in PowerPoint presentations
```java
// Create SmartArt shape
ISmartArt sa = ppt.getSlides().get(0).getShapes().appendSmartArt(200, 60, 300, 300, SmartArtLayoutType.GEAR);

// Set type and color of SmartArt
sa.setStyle(SmartArtStyleType.SUBTLE_EFFECT);
sa.setColorStyle(SmartArtColorType.GRADIENT_LOOP_ACCENT_3);

// Remove all shapes
for (Object a : sa.getNodes()) {
    sa.getNodes().removeNode(0);
}

// Add two custom shapes with text
ISmartArtNode node = sa.getNodes().addNode();
sa.getNodes().get(0).getTextFrame().setText("aa");
node = sa.getNodes().addNode();
node.getTextFrame().setText("bb");
node.getTextFrame().getTextRange().getFill().setFillType(FillFormatType.SOLID);
node.getTextFrame().getTextRange().getFill().getSolidColor().setKnownColor(KnownColors.BLACK);
```

---

# Extract Text from SmartArt
## Extract text content from SmartArt shapes in PowerPoint presentations
```java
//Traverse through all the slides of the PPT file and find the SmartArt shapes.
for (int i = 0; i < presentation.getSlides().getCount(); i++)
{
    for (int j = 0; j < presentation.getSlides().get(i).getShapes().getCount(); j++)
    {
        if (presentation.getSlides().get(i).getShapes().get(j) instanceof ISmartArt)
        {
            ISmartArt smartArt = (ISmartArt)presentation.getSlides().get(i).getShapes().get(j);

            //Extract text from SmartArt
            for (int k = 0; k < smartArt.getNodes().getCount(); k++)
            {
                String text = smartArt.getNodes().get(k).getTextFrame().getText();
            }
        }
    }
}
```

---

# SmartArt Node Removal
## Remove a specific node from SmartArt in a PowerPoint presentation
```java
// Get the SmartArt and collect nodes
ISmartArt sa = (ISmartArt) ((ppt.getSlides().get(0).getShapes().get(0) instanceof ISmartArt) ? ppt.getSlides().get(0).getShapes().get(0) : null);
ISmartArtNodeCollection nodes = sa.getNodes();

// Remove the node to specific position
nodes.removeNodeByPosition(2);
```

---

# Spire.Presentation Image Watermark
## Add an image as a watermark to a PowerPoint slide
```java
//Set the properties of SlideBackground, and then fill the image as watermark.
presentation.getSlides().get(0).getSlideBackground().setType(BackgroundType.CUSTOM);
presentation.getSlides().get(0).getSlideBackground().getFill().setFillType(FillFormatType.PICTURE);
presentation.getSlides().get(0).getSlideBackground().getFill().getPictureFill().setFillType(PictureFillType.STRETCH);
presentation.getSlides().get(0).getSlideBackground().getFill().getPictureFill().getPicture().setEmbedImage(image);
```

---

# Spire.Presentation Watermark
## Add watermark to PowerPoint presentation
```java
//Set the width and height of watermark string
int width= 400;
int height= 300;
//Define a rectangle range
Rectangle2D.Double rect = new Rectangle2D.Double((presentation.getSlideSize().getSize().getWidth() - width) / 2,
        (presentation.getSlideSize().getSize().getHeight() - height) / 2, width, height);

//Add a rectangle shape with a defined range
IAutoShape shape = presentation.getSlides().get(0).getShapes().appendShape(ShapeType.RECTANGLE, rect);

//Set the style of shape
shape.getFill().setFillType(FillFormatType.NONE);
shape.getShapeStyle().getLineColor().setColor(Color.white);
shape.setRotation(-45);
shape.getLocking().setSelectionProtection(true);
shape.getLine().setFillType(FillFormatType.NONE);

//Add text to shape
shape.getTextFrame().setText("E-iceblue");
PortionEx textRange = shape.getTextFrame().getTextRange();

//Set the style of the text range
textRange.getFill().setFillType(FillFormatType.SOLID);
textRange.getFill().getSolidColor().setColor(Color.pink);
textRange.setFontHeight(50);
```

---

# Remove Watermark from Presentation
## Code to remove text and image watermarks from PowerPoint slides
```java
// Remove text watermark by removing the shape which contains the text string "E-iceblue".
for (int i = 0; i < presentation.getSlides().getCount(); i++)
{
    for (int j = 0; j < presentation.getSlides().get(i).getShapes().getCount(); j++)
    {
        if (presentation.getSlides().get(i).getShapes().get(j) instanceof IAutoShape)
        {
            IAutoShape shape = (IAutoShape)presentation.getSlides().get(i).getShapes().get(j);
            if (shape.getTextFrame().getText().contains("E-iceblue"))
            {
                presentation.getSlides().get(i).getShapes().remove(shape);
            }
        }
    }
}

// Remove image watermark.
for (int i = 0; i < presentation.getSlides().getCount(); i++)
{
    presentation.getSlides().get(i).getSlideBackground().getFill().setFillType(FillFormatType.NONE);
}
```

---

# Spire Presentation OLE Object Embedding
## Embed Excel file as OLE object in presentation
```java
// Create a Presentation document
Presentation ppt = new Presentation();

// Insert an OLE object to presentation based on the Excel data
com.spire.presentation.IOleObject oleObject = ppt.getSlides().get(0).getShapes().appendOleObject("excel", data, rec);
oleObject.getSubstituteImagePictureFillFormat().getPicture().setEmbedImage(oleImage);
oleObject.setProgId("Excel.Sheet.12");
```

---

# Extract OLE Objects from PowerPoint Presentation
## This code extracts OLE (Object Linking and Embedding) objects from a PowerPoint presentation and identifies their type (Excel, Word, or PowerPoint documents).

```java
//Loop through the slides and shapes
for (Object slideObj : presentation.getSlides())
{
    ISlide slide=(ISlide)slideObj;
    for (Object shapeObj : slide.getShapes())
    {
        IShape shape=(IShape)shapeObj;
        if (shape instanceof IOleObject)
        {
            //Find OLE object
            IOleObject oleObject = (IOleObject)shape;

            //Get its data
            byte[] bytes = oleObject.getData();
            
            //Identify OLE object type
            switch (oleObject.getProgId())
            {
                case "Excel.Sheet.8":
                    //Excel .xls file
                    break;
                case "Excel.Sheet.12":
                    //Excel .xlsx file
                    break;
                case "Word.Document.8":
                    //Word .doc file
                    break;
                case "Word.Document.12":
                    //Word .docx file
                    break;
                case "PowerPoint.Show.8":
                    //PowerPoint .ppt file
                    break;
                case "PowerPoint.Show.12":
                    //PowerPoint .pptx file
                    break;
            }
        }
    }
};
```

---

# Get OLE Properties from PowerPoint
## Extract frame properties (height, width, top, left) from OLE objects in a PowerPoint presentation
```java
//create a PPT document
Presentation ppt = new Presentation();

//load ppt file
ppt.loadFromFile("data/GetOLEPropertiesOutsideOfShape.pptx");

//get the first slide
OleObjectCollection oles = ppt.getSlides().get(0).getOleObjects();

//get the first OLE
OleObject oleO = oles.get(0);

//get the information of OLE Object
oleO.getFrame().getHeight();
oleO.getFrame().getWidth();
oleO.getFrame().getTop();
oleO.getFrame().getLeft();
```

---

# Modify OLE Data in PowerPoint Presentation
## This code demonstrates how to find and modify OLE object data within a PowerPoint presentation
```java
//Loop through the slides and shapes
for (Object slideObj : presentation.getSlides())
{
    ISlide slide=(ISlide)slideObj;
    for (Object shapeObj : slide.getShapes())
    {
        IShape shape=(IShape)shapeObj;
        if (shape instanceof IOleObject)
        {
            //Find OLE object
            IOleObject oleObject = (IOleObject)shape;

            //Get its data
            byte[] bytes = oleObject.getData();
            ByteArrayInputStream pptStream = new ByteArrayInputStream(bytes);
            ByteArrayOutputStream stream = new ByteArrayOutputStream();
            if (oleObject.getProgId().equals("PowerPoint.Show.12"))
            {
                //Load the PPT stream
                Presentation ppt = new Presentation();
                ppt.loadFromStream(pptStream, FileFormat.AUTO);

                //Append an image in slide
                ppt.getSlides().get(0).getShapes().appendEmbedImage(ShapeType.RECTANGLE, "data/Logo.png", new Rectangle(50, 50, 100, 100));
                ppt.saveToFile(stream, FileFormat.PPTX_2013);

                //Modify the data
                oleObject.setData(stream.toByteArray());
            }
        }
    }
}
```

---

# Spire Presentation Print
## Print a PowerPoint presentation
```java
// Create a presentation object
Presentation ppt = new Presentation();

// Load the file
ppt.loadFromFile("data/print.pptx");
PresentationPrintDocument document = new PresentationPrintDocument(ppt);

// Print the file
document.print();
ppt.dispose();
```

---

# Print Multiple Slides Into One Page
## This example demonstrates how to print multiple PowerPoint slides on a single page using Spire.Presentation for Java.
```java
//Create a PPT document
Presentation ppt = new Presentation();

//Load the document from disk
ppt.loadFromFile("data/PrintMultipleSlidesIntoOnePage.pptx");
PresentationPrintDocument document = new PresentationPrintDocument(ppt);

//Set print task name
document.setDocumentName("print task 1");
document.setPrintOrder(Order.Horizontal);
document.setSlideFrameForPrint(true);

//Set Gray level when printing
document.setGrayLevelForPrint(true);
//Set four slides on one page
document.setSlideCountPerPageForPrint(PageSlideCount.Four);

//Set continuous print area
document.getPrinterSettings().setPrintRange(PrintRange.SomePages);
document.getPrinterSettings().setFromPage(1);
document.getPrinterSettings().setToPage(ppt.getSlides().getCount()-1);

ppt.print(document);
ppt.dispose();
```

---

# Print PowerPoint Presentation
## Print PPT file using virtual printer
```java
//Create a PowerPoint document.
Presentation presentation = new Presentation();

//Load the file from disk.
presentation.loadFromFile("data/Template_Ppt_6.pptx");

//Print PowerPoint document to virtual printer (Microsoft XPS Document Writer).
PresentationPrintDocument document = new PresentationPrintDocument(presentation);
document.getPrinterSettings().setPrinterName("Microsoft XPS Document Writer");

presentation.print(document);
```

---

# Print Specific PowerPoint Pages
## Code to print a specified range of pages from a PowerPoint presentation
```java
//Create a PowerPoint document.
Presentation presentation = new Presentation();

PresentationPrintDocument document = new PresentationPrintDocument(presentation);

//Set the document name to display while printing the document.
document.setDocumentName("Template_Ppt_6.pptx");

//Choose to print some pages from the PowerPoint document.
document.getPrinterSettings().setPrintRange(PrintRange.SomePages);
document.getPrinterSettings().setFromPage(2);
document.getPrinterSettings().setToPage(3);

short copyies=2;
//Set the number of copies of the document to print.
document.getPrinterSettings().setCopies(copyies);

presentation.print(document);
```

---

# Java PowerPoint Printing with Dialog
## Print PowerPoint presentation using print dialog
```java
//Create a PowerPoint document.
Presentation ppt = new Presentation();

//Get printer Job
PrinterJob printerJob= PrinterJob.getPrinterJob();
printerJob.setPrintable(ppt);
printerJob.printDialog();

//Print PPT
printerJob.print();
ppt.dispose();
```

---

# Spire Presentation Print Settings
## Configure print settings for PowerPoint presentation
```java
//Create a PowerPoint document.
Presentation presentation = new Presentation();

//Use PrintDocument object to print presentation slides.
PresentationPrintDocument document = new PresentationPrintDocument(presentation);

//Print document to virtual printer.
document.getPrinterSettings().setPrinterName("Microsoft XPS Document Writer");

//Print the slide with frame.
presentation.setSlideFrameForPrint(true);

//Print 4 slides horizontal.
presentation.setSlideCountPerPageForPrint(PageSlideCount.Four);
presentation.setOrderForPrint(Order.Horizontal);

//Print the slide with Grayscale.
presentation.setGrayLevelForPrint(true);

//Set the print document name.
document.setDocumentName("Template_Ppt_6.pptx");

document.getPrinterSettings().setPrintToFile(true);
document.getPrinterSettings().setPrintFileName("output/setPrintSettingsByPrintDocument.xps");

//Print the file
presentation.print(document);
```

---

# PowerPoint Print Settings Configuration
## Configure print settings for PowerPoint presentations using PrinterSettings
```java
// Create a PowerPoint document.
Presentation presentation = new Presentation();

// Use PrinterSettings object to print presentation slides.
PrinterSettings ps = new PrinterSettings();
ps.setPrintRange(PrintRange.AllPages);
ps.setPrintToFile(true);
String result = "output/setPrintSettingsByPrinterSettings.xps";
ps.setPrintFileName(result);

// Print the slide with frame.
presentation.setSlideFrameForPrint(true);

// Print the slide with Grayscale.
presentation.setGrayLevelForPrint(true);

// Print 4 slides horizontal.
presentation.setSlideCountPerPageForPrint(PageSlideCount.Four);
presentation.setOrderForPrint(Order.Horizontal);

// Only select some slides to print.
presentation.SelectSlidesForPrint("1", "3");

// Print the document.
presentation.print(ps);
```

---

# Spire Presentation Silent Printing
## Silently print PowerPoint presentation using default printer
```java
//Create a PowerPoint document.
Presentation presentation = new Presentation();

//Print the PowerPoint document to default printer.
PresentationPrintDocument document = new PresentationPrintDocument(presentation);
document.setPrintController(new StandardPrintController());

presentation.print(document);
```

---

# Spire Presentation Specific Printer Printing
## Print presentation to a specific printer with custom settings
```java
//New PrintSeetings
PrinterSettings printerSettings = new PrinterSettings();

//Set landscape for page
printerSettings.getDefaultPageSettings().setLandscape(true);

//Specific the printer
printerSettings.setPrinterName("Microsoft XPS Document Writer");

//Print
presentation.print(printerSettings);
```

---

# Spire Presentation VBA Macro Removal
## Remove VBA macros from PowerPoint presentation
```java
//Remove macros
//Note, at present it only can work on macros in PPT file, has not supported for PPTM file yet.
presentation.deleteMacros();
```

---

# Get PowerPoint Section
## Retrieve section information from a PowerPoint presentation
```java
// Create a PPT document
Presentation presentation = new Presentation();

// Load the file from disk
presentation.loadFromFile("data/GetSection.pptx");

// Get section list
SectionList list = presentation.getSectionList();
String name = list.get(0).getName();
```

---

