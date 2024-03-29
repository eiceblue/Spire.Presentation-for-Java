import com.spire.presentation.*;
import com.spire.presentation.charts.*;
import java.awt.*;
import java.awt.geom.Rectangle2D;

public class createScatterChart {
    public static void main(String[] args) throws Exception {
        String output = "output/createScatterChart.pptx";
        String imageFile = "data/bg.png";

        //create a presentation
        Presentation presentation = new Presentation();

        //set background image
        Rectangle2D.Double rect2 = new  Rectangle2D.Double(0, 0, presentation.getSlideSize().getSize().getWidth(), presentation.getSlideSize().getSize().getHeight());
        presentation.getSlides().get(0).getShapes().appendEmbedImage(ShapeType.RECTANGLE, imageFile, rect2);
        presentation.getSlides().get(0).getShapes().get(0).getLine().getFillFormat().getSolidFillColor().setColor(Color.white);

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

        //save the document
        presentation.saveToFile(output, FileFormat.PPTX_2013);
    }
}
