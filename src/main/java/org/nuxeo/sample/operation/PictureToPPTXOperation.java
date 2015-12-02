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
 */

package org.nuxeo.sample.operation;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
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
import org.nuxeo.ecm.core.api.DocumentModelList;
import org.nuxeo.ecm.core.api.DocumentRef;
import org.nuxeo.ecm.core.api.NuxeoException;
import org.nuxeo.sample.pptx4j.PictureToPPTX;

/**
 * @author jfletcher
 */
@Operation(id = PictureToPPTXOperation.ID, category = Constants.CAT_CONVERSION, label = "Picture To PPTX", description = "Convert a list of Pictures to Powerpoint presentation.")
public class PictureToPPTXOperation {

	public static final Log log = LogFactory
			.getLog(PictureToPPTXOperation.class);

	public static final String ID = "PictureToPPTX";

    @OperationMethod
	public Blob run(DocumentModelList input) throws NuxeoException {

		Blob outputBlob = null;

		try {
			PictureToPPTX myTest = new PictureToPPTX();
			myTest.setDocumentList(input);
			outputBlob = myTest.convert();
		} catch (Exception e) {
			throw new NuxeoException(e.getMessage());
		}

        return outputBlob;
	}

}
