/*
 * (C) Copyright 2015 Nuxeo SA (http://nuxeo.com/) and contributors.
 *
 * All rights reserved. This program and the accompanying materials
 * are made available under the terms of the GNU Lesser General Public License
 * (LGPL) version 2.1 which accompanies this distribution, and is available at
 * http://www.gnu.org/licenses/lgpl.html
 * 
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
 * Lesser General Public License for more details.
 * 
 * Contributors:
 *   Josh Fletcher <jfletcher@nuxeo.com>
 * 
 * This example is based on org.pptx4j.samples.InsertPicture, part of docx4j.
 * docx4j is copyright 2010-11, Plutext Pty Ltd.
 */

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
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.DocumentModel;
import org.nuxeo.ecm.core.api.DocumentModelList;
import org.nuxeo.ecm.core.api.impl.blob.FileBlob;
import org.nuxeo.ecm.platform.picture.api.PictureView;
import org.nuxeo.ecm.platform.picture.api.adapters.MultiviewPicture;
import org.nuxeo.runtime.api.Framework;
import org.pptx4j.jaxb.Context;
import org.pptx4j.pml.Pic;

/**
 * @author jfletcher
 */
public class PicturesToPPTX {

    public static final Log log = LogFactory.getLog(PicturesToPPTX.class);

    private static final String TITLE_MEDIUM = "Medium";

    public static final String PICTURE_TYPE = "Picture";

    private static final int PIXEL_TO_EMU_MULTIPLIER = 12700; // I believe this
                                                              // assumes 72dpi

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

    private DocumentModelList pictureList = null;

    /**
     * Constructor assumes you want to actually create a PPTX :)
     * 
     * @throws Exception
     */
    public PicturesToPPTX() throws Exception {
        presentationMLPackage = PresentationMLPackage.createPackage();
        // Need references to these parts to create a slide
        // Please note that these parts *already exist* - they are
        // created by createPackage() above. See that method
        // for instruction on how to create and add a part.
        mainPresentationPart = (MainPresentationPart) presentationMLPackage.getParts()
                                                                           .getParts()
                                                                           .get(new PartName("/ppt/presentation.xml"));
        layoutPart = (SlideLayoutPart) presentationMLPackage.getParts()
                                                            .getParts()
                                                            .get(new PartName("/ppt/slideLayouts/slideLayout1.xml"));
    }

    /**
     * At this point I'm going to assume I have Pictures; this is not defensive.
     * 
     * @param input
     */
    public void setDocumentList(DocumentModelList input) {
        pictureList = input;
    }

    /**
     * The constructor could probably handle this, but it feels safer to have a method to call.
     * 
     * @return
     * @throws Exception
     */
    public Blob convert() throws Exception {

        // Loop over the list of pictures, create a slide for each one.
        for (DocumentModel doc : pictureList) {
            if (!(PICTURE_TYPE.equals(doc.getType()))) {
                log.warn("Only " + PICTURE_TYPE + " documents are supported.");
            } else {
                addSlide(doc);
            }
        }

        File outputFile = File.createTempFile("nxops-" + this.getClass().getSimpleName() + "-", ".pptx");
        Framework.trackFile(outputFile, outputFile);

        // All done: save it
        presentationMLPackage.save(outputFile);

        // Load the file for Nuxeo.
        FileBlob outputFileBlob = new FileBlob(outputFile);

        return outputFileBlob;
    }

    /**
     * Given a Picture document, add a new slide to the presentation.
     * 
     * @param input
     * @throws Exception
     */
    private void addSlide(DocumentModel input) throws Exception {

        // I'm using the "Medium" Picture view.
        MultiviewPicture multiviewPicture = input.getAdapter(MultiviewPicture.class);
        PictureView mediumView = multiviewPicture.getView(TITLE_MEDIUM);

        pictureWidth = mediumView.getWidth();
        pictureHeight = mediumView.getHeight();
        pictureBlob = mediumView.getBlob();
        pictureTitle = input.getTitle();

        File imageFile = pictureBlob.getFile();

        slideCount++;

        // OK, now we can create a slide
        SlidePart slidePart = PresentationMLPackage.createSlidePart(mainPresentationPart, layoutPart, new PartName(
                "/ppt/slides/slide" + slideCount + ".xml"));

        // Add image part
        BinaryPartAbstractImage imagePart = BinaryPartAbstractImage.createImagePart(presentationMLPackage, slidePart,
                imageFile);

        // Add p:pic to slide
        slidePart.getJaxbElement()
                 .getCSld()
                 .getSpTree()
                 .getSpOrGrpSpOrGraphicFrame()
                 .add(createPicture(imagePart.getSourceRelationship().getId()));
    }

    /**
     * Generates the XML for the picture within the slide.
     * 
     * @param relId
     * @return
     * @throws JAXBException
     */
    protected Object createPicture(String relId) throws JAXBException {

        // Create p:pic
        java.util.HashMap<String, String> mappings = new java.util.HashMap<String, String>();

        mappings.put("id1", Long.toString(slideCount));
        mappings.put("name", pictureTitle);
        mappings.put("descr", pictureTitle);
        mappings.put("rEmbedId", relId);
        mappings.put("offx", Long.toString((PRESENTATION_WIDTH_EMU / 2 - pictureWidth * PIXEL_TO_EMU_MULTIPLIER / 2)));
        mappings.put("offy", Long.toString((PRESENTATION_HEIGHT_EMU / 2 - pictureHeight * PIXEL_TO_EMU_MULTIPLIER / 2)));
        mappings.put("extcx", Long.toString(pictureWidth * PIXEL_TO_EMU_MULTIPLIER));
        mappings.put("extcy", Long.toString(pictureHeight * PIXEL_TO_EMU_MULTIPLIER));

        return org.docx4j.XmlUtils.unmarshallFromTemplate(SAMPLE_PICTURE, mappings, Context.jcPML, Pic.class);

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
            + "</a:prstGeom>" + "</p:spPr>" + "</p:pic>";

}
