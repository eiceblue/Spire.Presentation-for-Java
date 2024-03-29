import com.spire.presentation.*;
import com.spire.presentation.diagrams.*;

public class removeNode {
    public static void main(String[] args) throws Exception{
        String inputFile = "data/RemoveNode.pptx";
        String outputFile = "output/removeNode.pptx";
        Presentation ppt = new Presentation();
        ppt.loadFromFile(inputFile);

        //Get the SmartArt and collect nodes
        ISmartArt sa = (ISmartArt) ((ppt.getSlides().get(0).getShapes().get(0) instanceof ISmartArt) ? ppt.getSlides().get(0).getShapes().get(0) : null);
        ISmartArtNodeCollection nodes = sa.getNodes();

        //Remove the node to specific position
        nodes.removeNodeByPosition(2);

        //Save to file
        ppt.saveToFile(outputFile, FileFormat.PPTX_2013);
        ppt.dispose();
    }
}
