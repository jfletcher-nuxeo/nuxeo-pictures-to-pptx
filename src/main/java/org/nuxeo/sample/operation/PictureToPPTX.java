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

/**
 * @author jfletcher
 */
@Operation(id = PictureToPPTX.ID, category = Constants.CAT_CONVERSION, label = "Picture To PPTX", description = "Convert a list of pictures to Powerpoint presentation.")
public class PictureToPPTX {

	public static final String ID = "PictureToPPTX";
	public static final String PICTURE_TYPE = "Picture";

	@OperationMethod(collector = BlobCollector.class)
	public Blob run(DocumentModel input) throws NuxeoException {
		if (!(PICTURE_TYPE.equals(input.getType()))) {
			throw new NuxeoException("Operation works only with "
					+ PICTURE_TYPE + " document type.");
		}

		return null;
	}

}
