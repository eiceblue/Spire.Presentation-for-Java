import com.spire.presentation.*;
import com.spire.presentation.drawing.*;

import java.awt.*;

public class helloWorld {
    public static void main(String[] args) throws Exception {
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

        //Save the document
        String outputFile = "output/helloWorld.pptx";
        presentation.saveToFile(outputFile, FileFormat.PPTX_2013);
    }
}
