import glob
import os
import email

def extract(extension="eml"):
    """
    try to extract the attachments from all files in cwd
    will process all files with the specified extension, or eml by default
    """
    # ensure that an output dir exists
    outputdir = "output"
    os.path.exists(outputdir) or os.makedirs(outputdir)
    output_count = 0
    ok = True
    for emlfile in glob.iglob("*.%s" % extension):
        print("FileName : %s " % emlfile)
        try:
            with open(emlfile, "r") as f:
                msg = email.message_from_file(f)
                attachmentsList = list(msg.get_payload()[1:])
                # If no attachments are found, skip this file
                    
                if not attachmentsList:
                    print("No attachment found for file %s!" % f.name)
                    ok = False
                    continue
                
                for attachment in attachmentsList:
                    
                    if not type(attachment) == type('String'):
                        #print("ContentType: &s " & attachment.get_content_type())
                        #print("ObjectType: %s" % type(attachment))
                        output_filename = attachment.get_filename()
                        
                        #html occurs error 
                        if type(output_filename) == type(None):
                            continue
                        
                        output_name = emlfile + "_" + output_filename
                        output_name = output_name.replace(".eml","")
                        output_name = output_name.replace(" ","_")
                        
                        with open(os.path.join(outputdir,output_name), 'w') as of:
                            of.write(attachment.get_payload(decode=True))
                            output_count += 1
        # this should catch read and write errors
        except IOError:
            print("There was a problem with %s or its attachment!" % f.name)
            ok = False

    if not ok:
        print("%s files were written, but there were some problems." % output_count)
    else:
        print(
            "Done. %s CSVs written to the 'output' directory." % output_count)

if __name__ == "__main__":
    extract()


