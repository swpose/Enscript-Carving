class MainClass {
    typedef byte[] ByteArray;
    void Main(CaseClass c) {
        SystemClass::ClearConsole(1);
        EntryFileClass file();
        int count = 0;
        int flag = 0;
        for (ItemIteratorClass i(c); EntryClass en = i.GetNextEntry();) {
            //open name Unallocated Clusters
            if (en.Name() == "Unallocated Clusters") {
                if (file.Open(en, FileClass::SLACK)) {
                    Console.WriteLine("File " + file.Name() + " is open.");
                    ByteArray ba(en.LogicalSize());
                    LocalFileClass localFile();
                    // Read Binary all data
                    file.ReadBinary(ba);
                    int i=0;
                    int j=0;
                    int k=0;
                    while (i < en.LogicalSize()){
                        if (flag == 0){
                            i++;
                            }
                        else{
                            i=k;
                            k++;
                        }
                        
                        //Find DOCX, XLSX Signature
                        if (ba[i-1]=='\x50' && ba[i] == '\x4B' && ba[i+1]== '\x03' && ba[i+2]=='\x04' && ba[i+3]=='\x14' && ba[i+4]=='\x00' && ba[i+5]=='\x06' && ba[i+6]=='\x00'){
                            Console.WriteLine(i);
                            flag = 0;
                            count+=1;
                            for (j= i; j< en.LogicalSize(); j++){
                                // Validation Flag for escape
                                if (flag == 1)
                                    break;
                                // Find 'word/' sentence
                                if (ba[j]=='\x77' && ba[j+1]=='\x6F' && ba[j+2]=='\x72' && ba[j+3] == '\x64' && ba[j+4] == '\x2F'){
                                    for (k=j; k<en.LogicalSize(); k++){
                                        //Find Footer
                                        if (ba[k]=='\x50' && ba[k+1]=='\x4B' && ba[k+2]=='\x05' && ba[k+3]=='\x06'){
                                            int offset = k+22; // followed by 18 additional bytes
                                            localFile.Open("C:\\en_test\\" + count + ".docx", FileClass::WRITE);
                                            for(int l=i-1; l< offset; l++){ // File Write
                                                localFile.WriteBinary(ba[l]);
                                            }
                                        // Flag set for escape
                                        flag = 1;
                                        }
                                        if (flag == 1)
                                            break;
                                    }
                                    localFile.Close();
                                }
                                // Find 'xl/' sentence
                                else if(ba[j]=='\x78' && ba[j+1]=='\x6C' && ba[j+2]=='\x2F'){
                                    for (k=j; k<en.LogicalSize(); k++){
                                        //Find Footer
                                        if (ba[k]=='\x50' && ba[k+1]=='\x4B' && ba[k+2]=='\x05' && ba[k+3]=='\x06'){
                                            int offset = k+22;
                                            localFile.Open("C:\\en_test\\" + count + ".xlsx", FileClass::WRITE);
                                            for(int l=i-1; l< offset; l++){ // followed by 18 additional bytes
                                                localFile.WriteBinary(ba[l]); // File Write
                                            }
                                        flag = 1;
                                        }
                                        if (flag ==1)
                                            break;
                                    }
                                    localFile.Close();
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

