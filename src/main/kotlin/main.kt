package ru.cadmean


import com.google.auth.oauth2.GoogleCredentials
import com.google.cloud.firestore.Firestore
import com.google.firebase.FirebaseApp
import com.google.firebase.FirebaseOptions
import com.google.firebase.cloud.FirestoreClient
import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.FileOutputStream
import java.util.*
import kotlin.collections.HashMap
import kotlin.math.roundToInt


fun main(args: Array<String>) {

    val serviceAccount = FileInputStream("C:\\Users\\msi-pc\\Documents\\GitHub\\ecopoints-748f3-firebase-adminsdk-bya4v-d967b2179c.json")
    val credentials = GoogleCredentials.fromStream(serviceAccount)
    val options = FirebaseOptions.Builder()
        .setCredentials(credentials)
        .build()
    FirebaseApp.initializeApp(options)

    val db: Firestore = FirestoreClient.getFirestore()
    
    val docRef = db.collection("user")


//    val output = FileOutputStream("Экосборка tab polly.xlsx")
//    val xlWbOut = XSSFWorkbook()
//    val xlWsOut = xlWbOut.createSheet()

    val input = FileInputStream("Экосборка tab polly.xlsx")
    val xlWbIn = WorkbookFactory.create(input)
    val xlWsIn = xlWbIn.getSheetAt(0)

    var i = 2
    var c = 1

    var clientList = HashMap<String, HashMap<String, Any>>()



    while(i<=830) {
        if(xlWsIn.getRow(i)!=null) {
            if (xlWsIn.getRow(i).getCell(1).getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
                var number = (xlWsIn.getRow(i).getCell(1).numericCellValue).toString()
                number = number.replaceFirst("8", "+7")
                number = number.replaceFirst(".", "")
                number = number.replaceFirst("E10", "")
                if (number.length == 12) {
                    if(xlWsIn.getRow(i).getCell(2) != null) {
                        val name = xlWsIn.getRow(i).getCell(2).toString()
                        val ecoPoints = xlWsIn.getRow(i).getCell(6).numericCellValue


                        val data: HashMap<String, Any> = HashMap()
                        data["firstName"] = name
                        data["middleName"] = " "
                        data["lastName"] = " "
                        data["role"] = "user"
                        data["phoneNumber"] = number
                        data["ecoPoints"] = ecoPoints.roundToInt()
                        data["i"] = i
                        data["c"] = c
                        data["repeat"] = false



                        if(clientList.containsKey(number)){
                            var t= clientList[number]!!["ecoPoints"] as Int
                            clientList[number]!!["ecoPoints"] = t + ecoPoints.roundToInt()
                            clientList[number]!!["repeat"] = true
                            i++
                            continue
                        }

                        clientList[number]=data

                        c++




//                        input.close()
//
//                        xlWsOut.createRow(i).createCell(12).setCellValue("добавлено Fb")
//                        xlWbOut.write(output)
//                        output.close()
                    }
                }
            }
        }
//i447
        //220
        i++
    }


    for(client in clientList){

    }



}
