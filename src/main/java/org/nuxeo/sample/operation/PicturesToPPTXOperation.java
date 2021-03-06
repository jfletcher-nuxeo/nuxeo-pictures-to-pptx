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
import org.nuxeo.ecm.automation.core.annotations.Operation;
import org.nuxeo.ecm.automation.core.annotations.OperationMethod;
import org.nuxeo.ecm.automation.core.annotations.Param;
import org.nuxeo.ecm.core.api.Blob;
import org.nuxeo.ecm.core.api.DocumentModelList;
import org.nuxeo.ecm.core.api.NuxeoException;
import org.nuxeo.sample.pptx4j.PicturesToPPTX;

/**
 * @author jfletcher
 */
@Operation(id = PicturesToPPTXOperation.ID, category = Constants.CAT_CONVERSION, label = "Pictures To PPTX", description = "Convert a list of Pictures to Powerpoint presentation. If <code>targetFileName</code> is not provided, the output filename will be \"nxops-PicturesToPPTX-&lt;some_number&gt;.pptx\".")
public class PicturesToPPTXOperation {

    public static final Log log = LogFactory.getLog(PicturesToPPTXOperation.class);

    public static final String ID = "PicturesToPPTX";

    @Param(name = "targetFileName", required = false)
    protected String targetFileName = null;

    @OperationMethod
    public Blob run(DocumentModelList documentList) throws NuxeoException {

        Blob outputBlob = null;

        try {
            PicturesToPPTX myTest = new PicturesToPPTX();
            myTest.setDocumentList(documentList);
            outputBlob = myTest.convert();
        } catch (Exception e) {
            throw new NuxeoException(e.getMessage());
        }

        if (targetFileName != null) {
            outputBlob.setFilename(targetFileName);
        }

        return outputBlob;
    }

}
