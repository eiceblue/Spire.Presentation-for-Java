
import com.spire.presentation.*;
import com.spire.presentation.drawing.FillFormatType;

import java.awt.*;

public class fillParticularRowWithColor {
    public static void main(String[] args) throws Exception {

        //Create a PPT document
        Presentation presentation = new Presentation();

        //Load the file from disk.
        presentation.loadFromFile("data/Template_Ppt_1.pptx");

        ITable table = null;

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

        String result = "output/Result-fillParticularRowWithColor.pptx";

        //Save to file.
        presentation.saveToFile(result, FileFormat.PPTX_2013);
    }
}

