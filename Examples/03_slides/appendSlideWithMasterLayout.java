import com.spire.presentation.*;
import com.spire.presentation.collections.*;
import com.spire.presentation.drawing.*;

import java.awt.*;
import java.awt.geom.*;

public class appendSlideWithMasterLayout {
    public static void main(String[] args) throws Exception {
        Presentation presentation = new Presentation();
        presentation.loadFromFile("data/appendSlideWithMasterLayout.pptx");

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


        //Save the document
        String result = "output/appendSlideWithMasterLayout.pptx";
        presentation.saveToFile(result, FileFormat.PPTX_2013);
    }
}
