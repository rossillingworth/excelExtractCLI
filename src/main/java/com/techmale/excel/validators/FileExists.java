package com.techmale.excel.validators;

import com.beust.jcommander.IParameterValidator;
import com.beust.jcommander.ParameterException;

import java.io.File;

/**
 * Make sure a file exists
 */
public class FileExists  implements IParameterValidator {
    public void validate(String name, String value) throws ParameterException {
        File f = new File(value);
        if (!f.exists()) {
            throw new ParameterException("Parameter " + name + ": unable to find file (" + value +")");
        }
    }
}
