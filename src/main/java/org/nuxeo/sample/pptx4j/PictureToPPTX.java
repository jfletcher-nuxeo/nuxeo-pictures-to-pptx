package org.nuxeo.sample.pptx4j;

import java.io.File;

import javax.xml.bind.JAXBException;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.docx4j.openpackaging.packages.PresentationMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.PresentationML.MainPresentationPart;
import org.docx4j.openpackaging.parts.PresentationML.SlideLayoutPart;
import org.docx4j.openpackaging.parts.PresentationML.SlidePart;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.relationships.Relationship;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.DocumentModel;
import org.nuxeo.ecm.core.api.impl.blob.FileBlob;
import org.nuxeo.ecm.core.api.model.Property;
import org.nuxeo.ecm.platform.picture.api.PictureView;
import org.nuxeo.ecm.platform.picture.api.adapters.MultiviewPicture;
import org.nuxeo.sample.operation.PictureToPPTXOperation;
import org.pptx4j.jaxb.Context;
import org.pptx4j.pml.Pic;

/**
 * @author jfletcher
 */
public class PictureToPPTX {

	public static final Log log = LogFactory
			.getLog(PictureToPPTXOperation.class);

	private static final String TITLE_MEDIUM = "Medium";
	private static final int PIXEL_TO_EMU_MULTIPLIER = 12700;
	private static final int PRESENTATION_WIDTH_EMU = 9902952; // 10.83 in
	private static final int PRESENTATION_HEIGHT_EMU = 6858000; // 7.5 in

	private int pictureWidth = 0;
	private int pictureHeight = 0;
	private Blob pictureBlob = null;
	private String pictureTitle = null;

	private int slideCount = 0;

	private PresentationMLPackage presentationMLPackage;
	private MainPresentationPart mainPresentationPart;
	private SlideLayoutPart layoutPart;
	
	// Where will we save our new .pptx?
	private String outputfilepath = "/Users/jfletcher/Nuxeo/Projects/prospects/Fox/test/pptx-picture.pptx";

	public PictureToPPTX() throws Exception {
		presentationMLPackage = PresentationMLPackage
				.createPackage();
		// Need references to these parts to create a slide
		// Please note that these parts *already exist* - they are
		// created by createPackage() above. See that method
		// for instruction on how to create and add a part.
		mainPresentationPart = (MainPresentationPart) presentationMLPackage
				.getParts().getParts()
				.get(new PartName("/ppt/presentation.xml"));
		layoutPart = (SlideLayoutPart) presentationMLPackage
				.getParts().getParts()
				.get(new PartName("/ppt/slideLayouts/slideLayout1.xml"));
	}

	public Blob convert(DocumentModel input) throws Exception {

		MultiviewPicture multiviewPicture = input
				.getAdapter(MultiviewPicture.class);
		PictureView mediumView = multiviewPicture.getView(TITLE_MEDIUM);

		pictureWidth = mediumView.getWidth();
		pictureHeight = mediumView.getHeight();
		pictureBlob = mediumView.getBlob();
		pictureTitle = input.getTitle();

		File imageFile = pictureBlob.getFile();

		slideCount++;

		// OK, now we can create a slide
		SlidePart slidePart = PresentationMLPackage.createSlidePart(mainPresentationPart,
				layoutPart, new PartName("/ppt/slides/slide1.xml"));

		// Add image part
		File file = new File("/Users/jfletcher/Documents/import/bg-img-1.jpg");
		BinaryPartAbstractImage imagePart = BinaryPartAbstractImage
				.createImagePart(presentationMLPackage, slidePart, imageFile);

		// Add p:pic to slide
		slidePart.getJaxbElement().getCSld().getSpTree()
				.getSpOrGrpSpOrGraphicFrame()
				.add(createPicture(imagePart.getSourceRelationship().getId()));

		File outputFile = new java.io.File(outputfilepath);

		// All done: save it
		presentationMLPackage.save(outputFile);

		FileBlob outputFileBlob = new FileBlob(outputFile);

		System.out.println("\n\n done .. saved " + outputfilepath);

		return outputFileBlob;
	}

	protected Object createPicture(String relId) throws JAXBException {

		// Create p:pic
		java.util.HashMap<String, String> mappings = new java.util.HashMap<String, String>();

		mappings.put("id1", Long.toString(slideCount));
		mappings.put("name", pictureTitle);
		mappings.put("descr", pictureTitle);
		mappings.put("rEmbedId", relId);
		mappings.put(
				"offx",
				Long.toString((PRESENTATION_WIDTH_EMU / 2 - pictureWidth
						* PIXEL_TO_EMU_MULTIPLIER / 2)));
		mappings.put(
				"offy",
				Long.toString((PRESENTATION_HEIGHT_EMU / 2 - pictureHeight
						* PIXEL_TO_EMU_MULTIPLIER / 2)));
		mappings.put("extcx",
				Long.toString(pictureWidth * PIXEL_TO_EMU_MULTIPLIER));
		mappings.put("extcy",
				Long.toString(pictureHeight * PIXEL_TO_EMU_MULTIPLIER));

		return org.docx4j.XmlUtils.unmarshallFromTemplate(SAMPLE_PICTURE,
				mappings, Context.jcPML, Pic.class);

	}

	private static String SAMPLE_PICTURE = "<p:pic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"> "
			+ "<p:nvPicPr>"
			+ "<p:cNvPr id=\"${id1}\" name=\"${name}\" descr=\"${descr}\"/>"
			+ "<p:cNvPicPr>"
			+ "<a:picLocks noChangeAspect=\"1\"/>"
			+ "</p:cNvPicPr>"
			+ "<p:nvPr/>"
			+ "</p:nvPicPr>"
			+ "<p:blipFill>"
			+ "<a:blip r:embed=\"${rEmbedId}\" cstate=\"print\"/>"
			+ "<a:stretch>"
			+ "<a:fillRect/>"
			+ "</a:stretch>"
			+ "</p:blipFill>"
			+ "<p:spPr>"
			+ "<a:xfrm>"
			+ "<a:off x=\"${offx}\" y=\"${offy}\"/>"
			+ "<a:ext cx=\"${extcx}\" cy=\"${extcy}\"/>"
			+ "</a:xfrm>"
			+ "<a:prstGeom prst=\"rect\">"
			+ "<a:avLst/>"
			+ "</a:prstGeom>"
			+ "</p:spPr>" + "</p:pic>";

}
