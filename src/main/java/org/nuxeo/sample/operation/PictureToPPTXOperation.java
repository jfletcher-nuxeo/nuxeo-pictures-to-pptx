/**
 * 
 */

package org.nuxeo.sample.operation;

import org.nuxeo.ecm.automation.core.Constants;
import org.nuxeo.ecm.automation.core.annotations.Context;
import org.nuxeo.ecm.automation.core.annotations.Operation;
import org.nuxeo.ecm.automation.core.annotations.OperationMethod;
import org.nuxeo.ecm.automation.core.annotations.Param;
import org.nuxeo.ecm.automation.core.collectors.DocumentModelCollector;
import org.nuxeo.ecm.automation.core.collectors.BlobCollector;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.CoreSession;
import org.nuxeo.ecm.core.api.DocumentModel;
import org.nuxeo.ecm.core.api.DocumentRef;
import org.nuxeo.ecm.core.api.NuxeoException;
import org.nuxeo.ecm.core.api.impl.blob.FileBlob;
import org.nuxeo.sample.pptx4j.PictureToPPTX;
import org.pptx4j.samples.InsertPicture;

/**
 * @author jfletcher
 */
@Operation(id = PictureToPPTXOperation.ID, category = Constants.CAT_CONVERSION, label = "Picture To PPTX", description = "Convert a list of Pictures to Powerpoint presentation.")
public class PictureToPPTXOperation {

	public static final String ID = "PictureToPPTX";
	public static final String PICTURE_TYPE = "Picture";

	@OperationMethod(collector = BlobCollector.class)
	public Blob run(DocumentModel input) throws NuxeoException {
		if (!(PICTURE_TYPE.equals(input.getType()))) {
			throw new NuxeoException("Operation works only with "
					+ PICTURE_TYPE + " document type.");
		}

		PictureToPPTX myTest = new PictureToPPTX();

		Blob outputBlob = null;
		
		try {
			outputBlob = myTest.run(input);
		} catch (Exception e) {
			e.printStackTrace();
		}

		return outputBlob;
	}

}
