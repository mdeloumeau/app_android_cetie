/*
==============================================================
Script : DetailActivity - Consultation et gestion d'une affaire
Auteur : Deloumeau Maël
Date   : 30/10/2025

But :
- Afficher les détails d'une affaire spécifique (numAffaire)
- Gestion des photos associées : visualisation, capture, upload, suppression
- Gestion des documents PDF / Word : ouverture, édition locale, conversion Word -> PDF, upload
- Validation des documents (FP, PVEE, PVEA) et mise à jour sur SharePoint
- Export / déplacement du dossier validé vers un répertoire "Valide" sur OneDrive

Librairies principales :
- MSAL : authentification Microsoft (Azure AD)
- OkHttp : requêtes HTTP pour Microsoft Graph API
- Glide : affichage des aperçus photos
- Android SDK : UI (Button, TextView, LinearLayout), Toast, Intent, FileProvider
- JSON : gestion des états de validation
- ActivityResultContracts : capture photo et édition de fichiers

Étapes principales :
1) Initialisation de l'UI : boutons, conteneur de photos, listeners
2) Authentification MSAL et acquisition du token Graph API
3) Chargement ou création du fichier validation.json sur OneDrive
4) Initialisation des sous-dossiers Photos et PV
5) Récupération, affichage et gestion des photos
6) Capture et upload de nouvelles photos
7) Ouverture, modification et réupload des fichiers PDF / Word
8) Validation des documents et mise à jour de l'état dans validation.json
9) Conversion Word -> PDF et suppression des fichiers Word
10) Déplacement du dossier validé vers "/Essais/Valide" sur OneDrive
==============================================================
*/
package com.cetie.appli_tablette

// -------------------------------------------------------------------------
// IMPORTS
// -------------------------------------------------------------------------
// Import des librairies Android, MSAL, OkHttp, Glide, JSON et ActivityResultContracts
// Gestion UI : Button, TextView, LinearLayout, ImageView, Toast, AlertDialog, Intent
// Gestion fichiers : File, FileProvider, MediaStore
// Gestion des couleurs et des états UI : Color, ColorStateList
// Microsoft Graph API : OkHttp pour GET, PUT, DELETE, PATCH
// MSAL : ISingleAccountPublicClientApplication, AuthenticationCallback
// Glide : affichage des images distantes avec authentification
// JSON : JSONObject, JSONArray


import android.app.AlertDialog
import android.content.Context
import android.content.Intent
import android.content.pm.ActivityInfo
import android.net.Uri
import android.os.Bundle
import android.provider.MediaStore
import android.view.ViewGroup
import android.widget.*
import androidx.activity.OnBackPressedCallback
import androidx.activity.result.ActivityResultLauncher
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.core.content.FileProvider
import okhttp3.*
import okhttp3.MediaType.Companion.toMediaTypeOrNull
import okhttp3.RequestBody.Companion.asRequestBody
import okhttp3.RequestBody.Companion.toRequestBody
import org.json.JSONObject
import java.io.File
import java.io.IOException
import java.text.SimpleDateFormat
import java.util.*
import com.bumptech.glide.Glide
import com.bumptech.glide.load.model.GlideUrl
import com.bumptech.glide.load.model.LazyHeaders
import com.bumptech.glide.request.RequestOptions
import com.microsoft.identity.client.*
import com.microsoft.identity.client.exception.MsalException
import android.content.res.ColorStateList
import android.view.inputmethod.InputMethodManager
import android.widget.Button
import org.json.JSONArray
import androidx.core.graphics.toColorInt
import androidx.core.net.toUri


class DetailActivity : AppCompatActivity() {

    // -------------------------------------------------------------------------
    // Variables globales
    // -------------------------------------------------------------------------

    // Infos dossier et fichiers
    private lateinit var numAffaire: String           // Numéro de l'affaire
    private lateinit var nomClient : String           // Nom du client
    private lateinit var folderId: String             // ID du dossier principal sur OneDrive
    private lateinit var driveId: String              // ID du drive OneDrive
    private var pvFolderId: String? = null           // ID du sous-dossier PV
    private var photosFolderId: String? = null       // ID du sous-dossier Photos

    // Authentification Microsoft
    private lateinit var pca: ISingleAccountPublicClientApplication // Instance MSAL
    private lateinit var authAccessToken: String                        // Token Graph API
    private val authority = "https://login.microsoftonline.com/69798bd7-8897-4dfe-9f84-1a861de23af6"    // URL d’authentification

    // UI et fichiers temporaires
    private lateinit var textPhotoCount: TextView  // Affiche le nombre de photos
    private lateinit var photoContainer: LinearLayout // Conteneur dynamique pour les aperçus de photos
    private lateinit var photoFile: File           // Fichier temporaire pour capture photo
    private var pdfFileBeingEdited: File? = null   // PDF ouvert pour édition
    private var lastPdfModified: Long = 0          // Dernière modification du PDF

    // Validation JSON
    private var validationJson = JSONObject()      // Stocke l'état de validation (FP, PVEE, PVEA)

    // ActivityResultLauncher pour ouverture PDF/Word
    private lateinit var pdfEditLauncher: ActivityResultLauncher<Intent>

    // -------------------------------------------------------------------------
    // onCreate()
    // -------------------------------------------------------------------------
    // Initialisation de l’activité, récupération des infos depuis MainActivity
    // Initialisation de l’authentification MSAL
    // Initialisation UI : boutons, conteneur photos, listeners

    override fun onCreate(savedInstanceState: Bundle?) {

        val isTablet = resources.configuration.smallestScreenWidthDp >= 600

        requestedOrientation = if (isTablet)
            ActivityInfo.SCREEN_ORIENTATION_LANDSCAPE
        else
            ActivityInfo.SCREEN_ORIENTATION_PORTRAIT

        super.onCreate(savedInstanceState)

        setContentView(R.layout.activity_detail)

        // Récupération infos de MainActivity
        numAffaire = intent.getStringExtra("NUM_AFFAIRE")?.take(8) ?: ""
        nomClient = intent.getStringExtra("NOM_CLIENT") ?: ""
        folderId = intent.getStringExtra("FOLDER_ID") ?: ""
        driveId = intent.getStringExtra("DRIVE_ID") ?: ""

        val buttonFPok = findViewById<Button>(R.id.buttonFPok)
        val buttonPVEEok = findViewById<Button>(R.id.buttonPVEEok)
        val buttonPVEAok = findViewById<Button>(R.id.buttonPVEAok)

        if (folderId.isEmpty() || driveId.isEmpty()) {
            Toast.makeText(this, "Dossier introuvable", Toast.LENGTH_LONG).show()
            finish()
            return
        }

        val textNumAffaire = findViewById<TextView>(R.id.textNumAffaire)
        textNumAffaire.text = numAffaire

        PublicClientApplication.createSingleAccountPublicClientApplication(
            applicationContext,
            R.raw.msal_config,
            object : IPublicClientApplication.ISingleAccountApplicationCreatedListener {
                override fun onCreated(application: ISingleAccountPublicClientApplication) {
                    pca = application
                    acquireTokenForDetail(buttonFPok, buttonPVEEok, buttonPVEAok)
                }
                override fun onError(e: MsalException) {
                    Toast.makeText(this@DetailActivity,"Erreur lors de l'initialisation MSAL: $e", Toast.LENGTH_LONG).show()
                }
            }
        )

        textPhotoCount = findViewById(R.id.textPhotoCount)
        photoContainer = findViewById(R.id.photoContainer)

        pdfEditLauncher = registerForActivityResult(ActivityResultContracts.StartActivityForResult()) {
            Toast.makeText(this, "PDF enregistré", Toast.LENGTH_SHORT).show()
        }

        initPdfButtons()
        initValidationButton(this)
        findViewById<Button>(R.id.buttonPhoto).setOnClickListener { takePhoto() }

        buttonFPok.setOnClickListener { toggleOkButton(buttonFPok) }
        buttonPVEEok.setOnClickListener { toggleOkButton(buttonPVEEok) }
        buttonPVEAok.setOnClickListener { toggleOkButton(buttonPVEAok) }

        val callback = object : OnBackPressedCallback(true) {
            override fun handleOnBackPressed() {
                val intent = Intent(this@DetailActivity, MainActivity::class.java)
                intent.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK or Intent.FLAG_ACTIVITY_CLEAR_TASK)
                startActivity(intent)
                finish()
            }
        }
        onBackPressedDispatcher.addCallback(this, callback)
    }

    // -------------------------------------------------------------------------
    // MSAL / Authentification
    // -------------------------------------------------------------------------
    // Création de l’instance MSAL (Single Account)
    // Une fois créée, acquisition du token Graph API
    // onError -> affiche un toast si MSAL échoue
    private fun acquireTokenForDetail(buttonFPok: Button, buttonPVEEok: Button, buttonPVEAok: Button) {
        val scopes = arrayOf("Files.ReadWrite")
        pca.acquireTokenSilentAsync(scopes, authority, object: AuthenticationCallback {
            override fun onSuccess(result: IAuthenticationResult) {
                authAccessToken = result.accessToken
                // ⚠️ NE PAS appeler loadOrCreateValidationFile avant le token
                runOnUiThread {
                    loadOrCreateValidationFile { json ->
                        applyValidationColors(buttonFPok, json.optBoolean("FP", false))
                        applyValidationColors(buttonPVEEok, json.optBoolean("PVEE", false))
                        applyValidationColors(buttonPVEAok, json.optBoolean("PVEA", false))
                    }
                }
                initFolder()
            }
            override fun onError(ex: MsalException) {
                runOnUiThread {
                    Toast.makeText(this@DetailActivity, "$ex", Toast.LENGTH_LONG).show()
                }
            }
            override fun onCancel() {}
        })
    }


    private fun initFolder() {
        val client = OkHttpClient()
        val url = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$folderId/children"
        val request = Request.Builder()
            .url(url)
            .addHeader("Authorization", "Bearer $authAccessToken")
            .build()

        client.newCall(request).enqueue(object : Callback {
            override fun onFailure(call: Call, e: IOException) {
                runOnUiThread {
                    Toast.makeText(this@DetailActivity, "Erreur récupération fichiers", Toast.LENGTH_LONG).show()
                }
            }

            override fun onResponse(call: Call, response: Response) {
                val body = response.body?.string() ?: return
                val folders = JSONObject(body).getJSONArray("value")


                photosFolderId = null
                pvFolderId = null

                for (i in 0 until folders.length()) {
                    val item = folders.getJSONObject(i)
                    val name = item.getString("name").lowercase()
                    if (item.has("folder")) {
                        if (name == "photos") photosFolderId = item.getString("id")
                        if (name == "pv") pvFolderId = item.getString("id")
                    }
                }

                if (photosFolderId != null) {
                    fetchPhotosFromFolder(photosFolderId!!)
                } else {
                    runOnUiThread {
                        Toast.makeText(this@DetailActivity, "Dossier Photos introuvable", Toast.LENGTH_SHORT).show()
                    }
                }

            }
        })
    }

    private fun downloadFile(fileId: String, fileName: String, onComplete: (File) -> Unit) {
        val client = OkHttpClient()
        val url = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$fileId/content"
        val request = Request.Builder()
            .url(url)
            .addHeader("Authorization", "Bearer $authAccessToken")
            .build()

        client.newCall(request).enqueue(object : Callback {
            override fun onFailure(call: Call, e: IOException) {
                runOnUiThread {
                    Toast.makeText(this@DetailActivity, "Erreur téléchargement", Toast.LENGTH_LONG).show()
                }
            }

            override fun onResponse(call: Call, response: Response) {
                val tempFile = File(cacheDir, fileName)
                response.body?.byteStream()?.use { input ->
                    tempFile.outputStream().use { output ->
                        input.copyTo(output)
                    }
                }
                runOnUiThread { onComplete(tempFile) }
            }
        })
    }

    // -------------------------------------------------------------------------
    // Photos
    // -------------------------------------------------------------------------
    // fetchPhotosFromFolder() -> Récupération des images à partir du sous-dossier "Photos"
    // takePhoto() -> ouvre l’appareil photo et crée un fichier temporaire
    // photoLauncher -> callback après capture photo pour upload
    // getNextPhotoFileName() -> génère nom unique pour photo
    // refreshPhotoButtons() -> affiche les aperçus photos et bouton suppression
    // uploadPhoto() -> envoie la photo vers OneDrive

    private fun fetchPhotosFromFolder(photosFolderId: String) {
        val client = OkHttpClient()
        val url = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$photosFolderId/children"
        val request = Request.Builder()
            .url(url)
            .addHeader("Authorization", "Bearer $authAccessToken")
            .build()

        client.newCall(request).enqueue(object : Callback {
            override fun onFailure(call: Call, e: IOException) {
                runOnUiThread {
                    Toast.makeText(this@DetailActivity, "Erreur récupération Photos", Toast.LENGTH_LONG).show()
                }
            }

            override fun onResponse(call: Call, response: Response) {
                val body = response.body?.string() ?: return
                val filesJson = JSONObject(body).getJSONArray("value")
                val photos = mutableListOf<JSONObject>()

                for (i in 0 until filesJson.length()) {
                    val item = filesJson.getJSONObject(i)
                    val name = item.getString("name").lowercase()
                    if (name.endsWith(".jpg") || name.endsWith(".png") || name.endsWith(".jpeg")) photos.add(item)
                }
                runOnUiThread { refreshPhotoButtons(photos) }
            }
        })
    }


    private fun uploadPhoto(localFile: File, remoteFileName: String) {
        if (!localFile.exists()) {
            Toast.makeText(this, "Fichier photo introuvable localement", Toast.LENGTH_LONG).show()
            return
        }

        val client = OkHttpClient()

        // Étape 1 : récupérer le dossier "Photos"
        val urlPhotosFolder = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$folderId/children"
        val request = Request.Builder()
            .url(urlPhotosFolder)
            .addHeader("Authorization", "Bearer $authAccessToken")
            .build()

        client.newCall(request).enqueue(object : Callback {
            override fun onFailure(call: Call, e: IOException) {
                runOnUiThread {
                    Toast.makeText(this@DetailActivity, "Erreur recherche dossier Photos", Toast.LENGTH_LONG).show()
                }
            }

            override fun onResponse(call: Call, response: Response) {
                val body = response.body?.string() ?: ""
                val folders = JSONObject(body).getJSONArray("value")

                for (i in 0 until folders.length()) {
                    val item = folders.getJSONObject(i)
                    if (item.has("folder") && item.getString("name").lowercase() == "photos") {
                        photosFolderId = item.getString("id")
                        break
                    }
                }

                if (photosFolderId == null) {
                    runOnUiThread {
                        Toast.makeText(this@DetailActivity, "Dossier Photos introuvable", Toast.LENGTH_LONG).show()
                    }
                    return
                }

                // Étape 2 : uploader dans le dossier Photos
                val encodedPhotoName = Uri.encode(remoteFileName)
                val uploadUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$photosFolderId:/$encodedPhotoName:/content"
                val requestBody = localFile.asRequestBody("image/jpeg".toMediaTypeOrNull())
                val uploadRequest = Request.Builder()
                    .url(uploadUrl)
                    .put(requestBody)
                    .addHeader("Authorization", "Bearer $authAccessToken")
                    .build()

                client.newCall(uploadRequest).enqueue(object : Callback {
                    override fun onFailure(call: Call, e: IOException) {
                        runOnUiThread {
                            Toast.makeText(this@DetailActivity, "Erreur upload photo : ${e.message}", Toast.LENGTH_LONG).show()
                        }
                    }

                    override fun onResponse(call: Call, response: Response) {
                        val responseBody = response.body?.string()
                        runOnUiThread {
                            if (response.isSuccessful) {
                                Toast.makeText(this@DetailActivity, "Photo enregistrée", Toast.LENGTH_SHORT).show()
                                initFolder() // Rafraîchir l’affichage
                            } else {
                                Toast.makeText(this@DetailActivity, "Erreur upload photo : ${response.code} - $responseBody", Toast.LENGTH_LONG).show()
                            }
                        }
                    }
                })
            }
        })
    }


    private fun takePhoto() {
        getNextPhotoFileName { nextFileName ->
            if (nextFileName == null) {
                Toast.makeText(this, "Impossible de générer le nom du fichier", Toast.LENGTH_SHORT).show()
                return@getNextPhotoFileName
            }

            try {
                photoFile = File(cacheDir, nextFileName)
                val photoUri: Uri = FileProvider.getUriForFile(
                    this,
                    "${packageName}.provider",
                    photoFile
                )

                val intent = Intent(MediaStore.ACTION_IMAGE_CAPTURE).apply {
                    putExtra(MediaStore.EXTRA_OUTPUT, photoUri)
                    addFlags(Intent.FLAG_GRANT_WRITE_URI_PERMISSION)
                    addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
                }

                photoLauncher.launch(intent)
            } catch (e: Exception) {
                e.printStackTrace()
                Toast.makeText(this, "Erreur création fichier : $e", Toast.LENGTH_SHORT).show()
            }
        }
    }


    private val photoLauncher =
        registerForActivityResult(ActivityResultContracts.StartActivityForResult()) { result ->
            if (result.resultCode == RESULT_OK) {
                runOnUiThread {
                    Toast.makeText(this@DetailActivity, "Chargement", Toast.LENGTH_SHORT).show()
                }
                uploadPhoto(photoFile, photoFile.name)
            } else {
                photoFile.delete()
                Toast.makeText(this, "Capture annulée", Toast.LENGTH_SHORT).show()
            }
        }

    private fun getNextPhotoFileName(onResult: (String?) -> Unit) {
        if (photosFolderId == null) {
            Toast.makeText(this, "Dossier Photos introuvable", Toast.LENGTH_SHORT).show()
            onResult(null)
            return
        }

        val client = OkHttpClient()
        val url = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$photosFolderId/children"
        val request = Request.Builder()
            .url(url)
            .addHeader("Authorization", "Bearer $authAccessToken")
            .build()

        client.newCall(request).enqueue(object : Callback {
            override fun onFailure(call: Call, e: IOException) {
                runOnUiThread {
                    Toast.makeText(this@DetailActivity, "Erreur récupération photos", Toast.LENGTH_LONG).show()
                }
                onResult(null)
            }

            override fun onResponse(call: Call, response: Response) {
                val body = response.body?.string() ?: return onResult(null)
                val filesJson = JSONObject(body).optJSONArray("value") ?: return onResult(null)

                val datePart = SimpleDateFormat("yyMMdd", Locale.FRANCE).format(Date())

                // Regex : PHOTO_yyMMdd_numAffaire_compteur.jpg
                val regex = Regex("^PHOTO_${datePart}_${numAffaire}_(\\d+)\\.jpg$", RegexOption.IGNORE_CASE)

                val existingCounters = mutableListOf<Int>()
                for (i in 0 until filesJson.length()) {
                    val name = filesJson.getJSONObject(i).getString("name")
                    val match = regex.find(name)
                    if (match != null) {
                        val compteur = match.groupValues[1].toIntOrNull()
                        if (compteur != null) existingCounters.add(compteur)
                    }
                }

                // Trouver le plus petit compteur manquant
                var nextCounter = 1
                while (existingCounters.contains(nextCounter)) nextCounter++

                val nextFileName = "PHOTO_${datePart}_${numAffaire}_${nextCounter}.jpg"
                runOnUiThread {
                    onResult(nextFileName)
                }
            }
        })
    }

    private fun refreshPhotoButtons(photos: List<JSONObject>) {
        photoContainer.removeAllViews()

        photos.forEach { fileObj ->
            val fileName = fileObj.getString("name")
            val fileId = fileObj.getString("id")

            val row = LinearLayout(this).apply {
                orientation = LinearLayout.HORIZONTAL
                layoutParams = LinearLayout.LayoutParams(
                    ViewGroup.LayoutParams.WRAP_CONTENT,
                    ViewGroup.LayoutParams.WRAP_CONTENT
                ).apply { setMargins(12, 12, 12, 12) }
            }

            // ImageView pour l'aperçu
            val imageView = ImageView(this).apply {
                layoutParams = LinearLayout.LayoutParams(200, 200)
                scaleType = ImageView.ScaleType.CENTER_CROP
                setOnClickListener {
                    runOnUiThread {
                        Toast.makeText(this@DetailActivity, "Chargement", Toast.LENGTH_SHORT).show()
                    }
                    downloadFile(fileId, fileName) { localFile ->
                        val uri: Uri = FileProvider.getUriForFile(
                            this@DetailActivity,
                            "${packageName}.provider",
                            localFile
                        )
                        val intent = Intent(Intent.ACTION_VIEW).apply {
                            setDataAndType(uri, "image/*")
                            addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
                        }
                        startActivity(intent)
                    }
                }
            }

            val glideUrl = GlideUrl(
                "https://graph.microsoft.com/v1.0/drives/$driveId/items/$fileId/content",
                LazyHeaders.Builder()
                    .addHeader("Authorization", "Bearer $authAccessToken")
                    .build()
            )

            if (!isDestroyed && !isFinishing) {
                Glide.with(this)
                    .load(glideUrl)
                    .apply(RequestOptions().placeholder(R.drawable.placeholder))
                    .centerCrop()
                    .into(imageView)
            }


            // Bouton suppression
            val deleteLabel = TextView(this).apply {
                text = "❌"
                textSize = 18f
                layoutParams = LinearLayout.LayoutParams(
                    ViewGroup.LayoutParams.WRAP_CONTENT,
                    ViewGroup.LayoutParams.WRAP_CONTENT
                ).apply { setMargins(12, 0, 0, 0) }

                setOnClickListener {
                    AlertDialog.Builder(this@DetailActivity)
                        .setTitle("Suppression")
                        .setMessage("Êtes-vous sûr de vouloir supprimer cette photo ?")
                        .setPositiveButton("Oui") { _, _ ->
                            val url = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$fileId"
                            val request = Request.Builder()
                                .url(url)
                                .delete()
                                .addHeader("Authorization", "Bearer $authAccessToken")
                                .build()
                            OkHttpClient().newCall(request).enqueue(object: Callback {
                                override fun onFailure(call: Call, e: IOException) {
                                    runOnUiThread {
                                        Toast.makeText(this@DetailActivity, "Erreur suppression", Toast.LENGTH_LONG).show()
                                    }
                                }
                                override fun onResponse(call: Call, response: Response) {
                                    runOnUiThread {
                                        initFolder()
                                        Toast.makeText(this@DetailActivity, "Photo supprimée", Toast.LENGTH_SHORT).show()
                                    }
                                }
                            })
                        }
                        .setNegativeButton("Non", null)
                        .show()
                }
            }

            row.addView(imageView)
            row.addView(deleteLabel)
            photoContainer.addView(row)
        }

        textPhotoCount.text = getString(R.string.photo_count, photos.size) //chaine de caractère sockée dans res/values/strings.xml
    }


    // -------------------------------------------------------------------------
    // PDF / Word
    // -------------------------------------------------------------------------
    // initPdfButtons() -> ouvre les PV et la Fiche de prod au format PDF ou Word
    // onResume() -> vérifie si PDF local a été modifié -> réupload sur OneDrive

    private fun initPdfButtons() {
        val buttonFP = findViewById<Button>(R.id.buttonFP)
        val buttonPVEE = findViewById<Button>(R.id.buttonPVEE)
        val buttonPVEA = findViewById<Button>(R.id.buttonPVEA)



        fun openFile(prefix: String) {
            runOnUiThread {
                Toast.makeText(this@DetailActivity, "Chargement", Toast.LENGTH_SHORT).show()
            }

            val client = OkHttpClient()
            val urlMain = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$folderId/children"
            val requestMain = Request.Builder()
                .url(urlMain)
                .addHeader("Authorization", "Bearer $authAccessToken")
                .build()

            client.newCall(requestMain).enqueue(object : Callback {
                override fun onFailure(call: Call, e: IOException) {
                    runOnUiThread {
                        Toast.makeText(this@DetailActivity, "Erreur accès dossier PV", Toast.LENGTH_SHORT).show()
                    }
                }

                override fun onResponse(call: Call, response: Response) {
                    val body = response.body?.string() ?: ""
                    val folders = JSONObject(body).getJSONArray("value")

                    if (pvFolderId == null) {
                        for (i in 0 until folders.length()) {
                            val item = folders.getJSONObject(i)
                            if (item.has("folder") && item.getString("name").lowercase() == "pv") {
                                pvFolderId = item.getString("id")
                                break
                            }
                        }
                    }

                    if (pvFolderId == null) {
                        runOnUiThread {
                            Toast.makeText(this@DetailActivity, "Sous-dossier PV introuvable", Toast.LENGTH_SHORT).show()
                        }
                        return
                    }

                    // Récupère la liste des fichiers dans /PV
                    val urlPdf = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$pvFolderId/children"
                    val requestPdf = Request.Builder()
                        .url(urlPdf)
                        .addHeader("Authorization", "Bearer $authAccessToken")
                        .build()

                    client.newCall(requestPdf).enqueue(object : Callback {
                        override fun onFailure(call: Call, e: IOException) {
                            runOnUiThread {
                                Toast.makeText(this@DetailActivity, "Erreur récupération fichiers", Toast.LENGTH_SHORT).show()
                            }
                        }

                        override fun onResponse(call: Call, response: Response) {
                            val bodyPdf = response.body?.string() ?: ""
                            val filesJson = JSONObject(bodyPdf).getJSONArray("value")

                            val fileObj = (0 until filesJson.length())
                                .map { filesJson.getJSONObject(it) }
                                .firstOrNull { it.getString("name").startsWith("${prefix}_${numAffaire}") }

                            if (fileObj == null) {
                                if (prefix.equals("PVEA", ignoreCase = true)) {
                                    // PVEA introuvable -> proposer sélection de PVEA standards
                                    runOnUiThread {
                                        showPVEASelectionAndCopy()
                                    }
                                    return
                                } else {
                                    runOnUiThread {
                                        Toast.makeText(this@DetailActivity, "$prefix non trouvé", Toast.LENGTH_SHORT).show()
                                    }
                                    return
                                }
                            }


                            val fileId = fileObj.getString("id")
                            val fileName = fileObj.getString("name")

                            when {
                                fileName.endsWith(".pdf", ignoreCase = true) -> {
                                    // Télécharger et ouvrir localement
                                    downloadFile(fileId, fileName) { localFile ->
                                        if (!localFile.exists()) {
                                            runOnUiThread {
                                                Toast.makeText(this@DetailActivity, "Erreur fichier local PDF", Toast.LENGTH_SHORT).show()
                                            }
                                            return@downloadFile
                                        }

                                        pdfFileBeingEdited = localFile
                                        lastPdfModified = localFile.lastModified()

                                        val uri: Uri = FileProvider.getUriForFile(
                                            this@DetailActivity,
                                            "${packageName}.provider",
                                            localFile
                                        )

                                        val intent = Intent(Intent.ACTION_VIEW).apply {
                                            setDataAndType(uri, "application/pdf")
                                            addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION or Intent.FLAG_GRANT_WRITE_URI_PERMISSION)
                                        }

                                        if (intent.resolveActivity(packageManager) != null) {
                                            startActivity(intent)
                                        } else {
                                            runOnUiThread {
                                                Toast.makeText(this@DetailActivity, "Aucune app PDF installée (Xodo conseillée)", Toast.LENGTH_LONG).show()
                                            }
                                        }
                                    }
                                }

                                fileName.endsWith(".docx", ignoreCase = true) -> {
                                    // Ouvrir via OneDrive
                                    val linkRequestBody = """
                                { "type": "edit" }
                            """.trimIndent()

                                    val createLinkUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$fileId/createLink"
                                    val requestLink = Request.Builder()
                                        .url(createLinkUrl)
                                        .addHeader("Authorization", "Bearer $authAccessToken")
                                        .addHeader("Content-Type", "application/json")
                                        .post(linkRequestBody.toRequestBody("application/json".toMediaTypeOrNull()))
                                        .build()

                                    client.newCall(requestLink).enqueue(object : Callback {
                                        override fun onFailure(call: Call, e: IOException) {
                                            runOnUiThread {
                                                Toast.makeText(this@DetailActivity, "Erreur création lien OneDrive", Toast.LENGTH_SHORT).show()
                                            }
                                        }

                                        override fun onResponse(call: Call, response: Response) {
                                            val linkBody = response.body?.string() ?: ""
                                            val linkUrl = JSONObject(linkBody)
                                                .getJSONObject("link")
                                                .getString("webUrl")

                                            val intent = Intent(Intent.ACTION_VIEW, linkUrl.toUri())
                                            startActivity(intent)
                                        }
                                    })
                                }

                                fileName.endsWith(".xlsx", ignoreCase = true) || fileName.endsWith(".xls", ignoreCase = true) ||
                                        fileName.endsWith(".xlsm", ignoreCase = true) -> {

                                    // Ouvrir via OneDrive comme un fichier Word
                                    val linkRequestBody = """{ "type": "edit" }""".trimIndent()

                                    val createLinkUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$fileId/createLink"
                                    val requestLink = Request.Builder()
                                        .url(createLinkUrl)
                                        .addHeader("Authorization", "Bearer $authAccessToken")
                                        .addHeader("Content-Type", "application/json")
                                        .post(linkRequestBody.toRequestBody("application/json".toMediaTypeOrNull()))
                                        .build()

                                    client.newCall(requestLink).enqueue(object : Callback {
                                        override fun onFailure(call: Call, e: IOException) {
                                            runOnUiThread {
                                                Toast.makeText(this@DetailActivity, "Erreur création lien OneDrive", Toast.LENGTH_SHORT).show()
                                            }
                                        }

                                        override fun onResponse(call: Call, response: Response) {
                                            val linkBody = response.body?.string() ?: ""
                                            val linkUrl = JSONObject(linkBody)
                                                .getJSONObject("link")
                                                .getString("webUrl")

                                            val intent = Intent(Intent.ACTION_VIEW, linkUrl.toUri())
                                            startActivity(intent)
                                        }
                                    })
                                }


                                else -> runOnUiThread {
                                    Toast.makeText(this@DetailActivity, "Format de fichier non supporté", Toast.LENGTH_SHORT).show()
                                }
                            }
                        }
                    })
                }
            })
        }

        buttonFP.setOnClickListener { openFile("FP") }
        buttonPVEE.setOnClickListener { openFile("PVEE") }
        buttonPVEA.setOnClickListener { openFile("PVEA") }

    }



    override fun onResume() {
        super.onResume()
        pdfFileBeingEdited?.let { file ->
            val newModified = file.lastModified()
            if (newModified != lastPdfModified) {
                runOnUiThread {
                    Toast.makeText(this@DetailActivity, "Chargement", Toast.LENGTH_SHORT).show()
                }
                lastPdfModified = newModified

                if (pvFolderId == null) {
                    Toast.makeText(this, "Erreur : dossier PV introuvable pour upload", Toast.LENGTH_LONG).show()
                    return
                }

                if (!file.exists()) {
                    Toast.makeText(this, "Erreur : fichier PDF introuvable localement", Toast.LENGTH_LONG).show()
                    return
                }

                val encodedName = Uri.encode(file.name)
                val uploadUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$pvFolderId:/$encodedName:/content"
                val requestBody = file.asRequestBody("application/pdf".toMediaTypeOrNull())
                val request = Request.Builder()
                    .url(uploadUrl)
                    .put(requestBody)
                    .addHeader("Authorization", "Bearer $authAccessToken")
                    .build()

                OkHttpClient().newCall(request).enqueue(object : Callback {
                    override fun onFailure(call: Call, e: IOException) {
                        runOnUiThread {
                            Toast.makeText(this@DetailActivity, "Erreur réupload PDF : ${e.message}", Toast.LENGTH_LONG).show()
                        }
                    }

                    override fun onResponse(call: Call, response: Response) {
                        val responseBody = response.body?.string()
                        runOnUiThread {
                            if (response.isSuccessful) {
                                Toast.makeText(this@DetailActivity, "PDF mis à jour sur OneDrive", Toast.LENGTH_SHORT).show()
                            } else {
                                Toast.makeText(this@DetailActivity, "Erreur réupload PDF : ${response.code} - $responseBody", Toast.LENGTH_LONG).show()
                            }
                        }
                    }
                })
            }
            pdfFileBeingEdited = null
        }
    }

    // ------------------------
    // PVEA Standards : listing, dialogue de recherche, copie/rename
    // ------------------------

    data class StandardFile(val id: String, val name: String) // name inclut l'extension

    private fun showPVEASelectionAndCopy() {
        // 1) Vérifier qu'on a le driveId
        if (driveId.isBlank()) {
            Toast.makeText(this, "Drive introuvable", Toast.LENGTH_LONG).show()
            return
        }

        // Chemin relatif demandé : Production - Documents/1-Essais/4-PVEA standards
        val standardsPath = "1-Essais/4-PVEA standards"
        val url = "https://graph.microsoft.com/v1.0/drives/$driveId/root:/$standardsPath:/children"

        val client = OkHttpClient()
        val request = Request.Builder()
            .url(url)
            .addHeader("Authorization", "Bearer $authAccessToken")
            .build()

        client.newCall(request).enqueue(object : Callback {
            override fun onFailure(call: Call, e: IOException) {
                runOnUiThread {
                    Toast.makeText(this@DetailActivity, "Erreur récupération PVEA standards", Toast.LENGTH_LONG).show()
                }
            }

            override fun onResponse(call: Call, response: Response) {
                val bodyStr = response.body?.string() ?: ""
                if (!response.isSuccessful) {
                    runOnUiThread {
                        Toast.makeText(this@DetailActivity, "Erreur récupération PVEA standards : ${response.code}", Toast.LENGTH_LONG).show()
                    }
                    return
                }

                val items = JSONObject(bodyStr).optJSONArray("value") ?: return runOnUiThread {
                    Toast.makeText(this@DetailActivity, "Aucun PVEA standard trouvé", Toast.LENGTH_SHORT).show()
                }

                val standards = mutableListOf<StandardFile>()
                for (i in 0 until items.length()) {
                    val it = items.getJSONObject(i)
                    val name = it.optString("name")
                    // Ne garder que les Word (selon ta remarque : il n'y aura que des .docx)
                    if (name.endsWith(".docx", ignoreCase = true) || name.endsWith(".pdf", ignoreCase = true)) {
                        standards.add(StandardFile(it.optString("id"), name))
                    }
                }

                if (standards.isEmpty()) {
                    runOnUiThread {
                        Toast.makeText(this@DetailActivity, "Aucun PVEA standard disponible", Toast.LENGTH_SHORT).show()
                    }
                    return
                }

                // Construire la liste de display (sans extension) et afficher la popup (UI thread)
                runOnUiThread {
                    showPVEASelectionDialog(standards)
                }
            }
        })
    }

    /**
     * Affiche un AlertDialog contenant :
     * - EditText (search)
     * - ListView (liste filtrée)
     * L'affichage montre le nom SANS extension
     */
    private fun showPVEASelectionDialog(standards: List<StandardFile>) {
        // Copies mutables pour le filtre
        val allDisplayNames = standards.map { it.name.substringBeforeLast('.') } // sans extension

        // Builder dialog
        val builder = AlertDialog.Builder(this)
        builder.setTitle("Sélectionner un PVEA standard")

        // Layout simple : Linear vertical avec EditText puis ListView
        val container = LinearLayout(this).apply {
            orientation = LinearLayout.VERTICAL
            setPadding(20, 20, 20, 0)
        }

        val search = EditText(this).apply {
            hint = "Rechercher..."
            isSingleLine = true
        }

        val listView = ListView(this)
        val adapter = ArrayAdapter(this, android.R.layout.simple_list_item_single_choice, allDisplayNames.toMutableList())
        listView.choiceMode = ListView.CHOICE_MODE_SINGLE
        listView.adapter = adapter

        container.addView(search)
        container.addView(listView)
        builder.setView(container)

        // Buttons
        builder.setNegativeButton("Annuler") { dialog, _ ->
            dialog.dismiss() // juste fermer
        }

        builder.setPositiveButton("Insérer", null) // on override pour contrôle activation

        val dialog = builder.create()
        dialog.setOnShowListener {
            val insertButton = dialog.getButton(AlertDialog.BUTTON_POSITIVE)
            insertButton.isEnabled = false

            // Activation insert seulement si une sélection
            listView.setOnItemClickListener { _, _, position, _ ->
                insertButton.isEnabled = true
            }

            // Filtre en tapant
            search.addTextChangedListener(object : android.text.TextWatcher {
                override fun beforeTextChanged(s: CharSequence?, start: Int, count: Int, after: Int) {}
                override fun afterTextChanged(s: android.text.Editable?) {}
                override fun onTextChanged(s: CharSequence?, start: Int, before: Int, count: Int) {
                    val q = s?.toString()?.trim()?.lowercase() ?: ""
                    val filtered = if (q.isEmpty()) allDisplayNames else {
                        allDisplayNames.filter { it.lowercase().contains(q) }
                    }
                    // Mettre à jour adapter
                    adapter.clear()
                    adapter.addAll(filtered)
                    adapter.notifyDataSetChanged()
                    insertButton.isEnabled = false
                }
            })

            search.setOnEditorActionListener { v, actionId, event ->
                if (actionId == android.view.inputmethod.EditorInfo.IME_ACTION_DONE) {
                    // Cacher le clavier
                    val imm = getSystemService(Context.INPUT_METHOD_SERVICE) as InputMethodManager
                    imm.hideSoftInputFromWindow(v.windowToken, 0)
                    true // on consomme l'action
                } else {
                    false
                }
            }

            insertButton.setOnClickListener {
                val checkedPos = listView.checkedItemPosition
                if (checkedPos == ListView.INVALID_POSITION) return@setOnClickListener

                // Trouver le StandardFile correspondant au display name sélectionné
                val selectedDisplay = adapter.getItem(checkedPos) ?: return@setOnClickListener

                // Recherche index dans la liste originale (les noms sans extension)
                val originalIndex = allDisplayNames.indexOf(selectedDisplay)
                if (originalIndex < 0) {
                    // S'il n'est pas trouvé (rare, si filtrage), chercher par égalité de string en list
                    val matchIndex = allDisplayNames.indexOfFirst { it == selectedDisplay }
                    if (matchIndex < 0) {
                        Toast.makeText(this, "Fichier introuvable", Toast.LENGTH_SHORT).show()
                        dialog.dismiss()
                        return@setOnClickListener
                    } else {
                        copySelectedStandard(standards[matchIndex])
                    }
                } else {
                    copySelectedStandard(standards[originalIndex])
                }
                dialog.dismiss()
            }
        }

        dialog.show()
    }

    /**
     * Copie le fichier standard choisi dans le dossier PV de l'affaire et renomme.
     * - standardsFile : StandardFile (id + name)
     * - new name : PVEA_<NUM_AFFAIRE>_<CLIENT>.<extension>
     */
    private fun copySelectedStandard(standardFile: StandardFile) {
        // Vérifications
        if (pvFolderId == null) {
            // Tenter de retrouver le dossier PV (comme dans initFolder)
            Toast.makeText(this, "Traitement en cours : récupération dossier PV", Toast.LENGTH_SHORT).show()
            // relancer initFolder et ensuite rappeler la copie : on va lister le dossier parent pour récupérer PV
            // Simplification : redemander children du dossier principal (folderId) pour trouver PV
            val client = OkHttpClient()
            val url = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$folderId/children"
            val request = Request.Builder()
                .url(url)
                .addHeader("Authorization", "Bearer $authAccessToken")
                .build()

            client.newCall(request).enqueue(object : Callback {
                override fun onFailure(call: Call, e: IOException) {
                    runOnUiThread {
                        Toast.makeText(this@DetailActivity, "Erreur récupération dossier PV", Toast.LENGTH_LONG).show()
                    }
                }

                override fun onResponse(call: Call, response: Response) {
                    val body = response.body?.string() ?: ""
                    val files = JSONObject(body).optJSONArray("value") ?: JSONArray()
                    for (i in 0 until files.length()) {
                        val item = files.getJSONObject(i)
                        if (item.has("folder") && item.getString("name").equals("PV", true)) {
                            pvFolderId = item.getString("id")
                            break
                        }
                    }
                    if (pvFolderId == null) {
                        runOnUiThread {
                            Toast.makeText(this@DetailActivity, "Aucun dossier PV trouvé dans l'affaire", Toast.LENGTH_LONG).show()
                        }
                        return
                    }
                    // retry copy now that pvFolderId is known
                    copySelectedStandard(standardFile)
                }
            })
            return
        }

        // Construire nouveau nom : PVEA_<NUM_AFFAIRE>_<CLIENT>.<ext>
        val clientName = nomClient.replace("\\s+".toRegex(), "_")
        val extension = standardFile.name.substringAfterLast('.', "") // docx ou pdf
        val newName = "PVEA_${numAffaire}_${clientName}.${extension}"

        // POST copy endpoint
        val copyUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/${standardFile.id}/copy"
        val json = JSONObject().apply {
            put("parentReference", JSONObject().put("id", pvFolderId))
            put("name", newName)
        }
        val body = json.toString().toRequestBody("application/json".toMediaTypeOrNull())

        val request = Request.Builder()
            .url(copyUrl)
            .post(body)
            .addHeader("Authorization", "Bearer $authAccessToken")
            .addHeader("Content-Type", "application/json")
            .build()

        val client = OkHttpClient()
        runOnUiThread {
            Toast.makeText(this, "Copie en cours...", Toast.LENGTH_SHORT).show()
        }

        client.newCall(request).enqueue(object : Callback {
            override fun onFailure(call: Call, e: IOException) {
                runOnUiThread {
                    Toast.makeText(this@DetailActivity, "Erreur copie PVEA : ${e.message}", Toast.LENGTH_LONG).show()
                }
            }

            override fun onResponse(call: Call, response: Response) {
                // Graph : copy retourne généralement 202 Accepted (opération asynchrone)
                runOnUiThread {
                    if (response.isSuccessful || response.code == 202) {
                        Toast.makeText(this@DetailActivity, "PVEA inséré dans PV", Toast.LENGTH_SHORT).show()
                        // Rafraîchir l'affichage des fichiers PV
                        // Si tu veux, relancer initFolder() pour recharger photos et pvFolderId etc.
                        initFolder()
                    } else {
                        val respBody = response.body?.string()
                        Toast.makeText(this@DetailActivity, "Erreur copie PVEA : ${response.code} - $respBody", Toast.LENGTH_LONG).show()
                    }
                }
            }
        })
    }


    // -------------------------------------------------------------------------
    // Validation des documents (Boutons OK)
    // -------------------------------------------------------------------------
    // toggleOkButton() -> change couleur bouton et met à jour JSON
    // loadOrCreateValidationFile() -> récupère ou crée validation.json
    // uploadValidationJson() -> sauvegarde sur OneDrive

    private fun toggleOkButton(button: Button) {
        val green = "#4CAF50".toColorInt() // Validé
        val grey = "#9E9E9E".toColorInt()  // Non validé
        val red = "#F44336".toColorInt()   // Non nécessaire

        val key = when (button.id) {
            R.id.buttonFPok -> "FP"
            R.id.buttonPVEEok -> "PVEE"
            R.id.buttonPVEAok -> "PVEA"
            else -> ""
        }

        if (key.isEmpty()) return

        if (button.id == R.id.buttonPVEAok) {
            // Cas spécial : menu à 3 états
            val options = arrayOf("Validé", "Non nécessaire", "Non validé")

            AlertDialog.Builder(this)
                .setTitle("État du PV Automatique")
                .setItems(options) { _, which ->
                    val (color, state) = when (which) {
                        0 -> green to "validé"
                        1 -> red to "non_necessaire"
                        else -> grey to "non_valide"
                    }

                    button.backgroundTintList = ColorStateList.valueOf(color)
                    validationJson.put("PVEA", state)
                    uploadValidationJson(validationJson)
                }
                .show()

        } else {
            // Cas normal : toggle simple (gris ↔ vert)
            val currentColor = button.backgroundTintList?.defaultColor
            val isCurrentlyGreen = (currentColor == green)
            val actionText =
                if (isCurrentlyGreen) "retirer la validation" else "valider ce document"

            AlertDialog.Builder(this)
                .setTitle("Confirmation")
                .setMessage("Voulez-vous $actionText ?")
                .setPositiveButton("Oui") { _, _ ->
                    val nextColor = if (isCurrentlyGreen) grey else green
                    button.backgroundTintList = ColorStateList.valueOf(nextColor)
                    val newValue = (nextColor == green)
                    validationJson.put(key, newValue)
                    uploadValidationJson(validationJson)
                }
                .setNegativeButton("Non", null)
                .show()
        }
    }

    private fun loadOrCreateValidationFile(onLoaded: (JSONObject) -> Unit) {
        val client = OkHttpClient()
        val url = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$folderId/children"
        val request = Request.Builder()
            .url(url)
            .addHeader("Authorization", "Bearer $authAccessToken")
            .build()

        client.newCall(request).enqueue(object : Callback {
            override fun onFailure(call: Call, e: IOException) {
                runOnUiThread {
                    Toast.makeText(this@DetailActivity, "Erreur récupération fichiers", Toast.LENGTH_LONG).show()
                }
            }

            override fun onResponse(call: Call, response: Response) {
                val body = response.body?.string() ?: ""
                val files = JSONObject(body).optJSONArray("value") ?: return
                var validationFileId: String? = null

                for (i in 0 until files.length()) {
                    val item = files.getJSONObject(i)
                    if (item.getString("name") == "validation.json") {
                        validationFileId = item.getString("id")
                        break
                    }
                }

                if (validationFileId != null) {
                    // Télécharger le fichier existant
                    downloadFile(validationFileId, "validation.json") { file ->
                        validationJson = JSONObject(file.readText())
                        runOnUiThread { onLoaded(validationJson) }
                    }
                } else {
                    // Créer un nouveau JSON avec tout false
                    validationJson = JSONObject().apply {
                        put("FP", false)
                        put("PVEE", false)
                        put("PVEA", "non_valide")
                    }
                    runOnUiThread { onLoaded(validationJson) }

                    // Envoyer sur OneDrive
                    uploadValidationJson(validationJson)
                }
            }
        })
    }

    private fun applyValidationColors(button: Button, isValidated: Boolean) {
        val green = "#4CAF50".toColorInt()
        val grey = "#9E9E9E".toColorInt()
        val red = "#F44336".toColorInt()

        if (button.id == R.id.buttonPVEAok) {
            val state = validationJson.optString("PVEA", "non_valide")
            val color = when (state) {
                "validé" -> green
                "non_necessaire" -> red
                else -> grey
            }
            button.backgroundTintList = ColorStateList.valueOf(color)
        } else {
            val color = if (isValidated) green else grey
            button.backgroundTintList = ColorStateList.valueOf(color)
        }
    }


    private fun uploadValidationJson(json: JSONObject) {
        val encodedName = Uri.encode("validation.json")
        val url = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$folderId:/$encodedName:/content"
        val body = json.toString().toRequestBody("application/json".toMediaTypeOrNull())
        val request = Request.Builder()
            .url(url)
            .put(body)
            .addHeader("Authorization", "Bearer $authAccessToken")
            .build()

        OkHttpClient().newCall(request).enqueue(object: Callback {
            override fun onFailure(call: Call, e: IOException) {
                runOnUiThread {
                    Toast.makeText(this@DetailActivity, "Erreur sauvegarde validation.json", Toast.LENGTH_LONG).show()
                }
            }

            override fun onResponse(call: Call, response: Response) {
                runOnUiThread {
                    if (!response.isSuccessful) {
                        Toast.makeText(this@DetailActivity, "Erreur sauvegarde validation.json: ${response.code}", Toast.LENGTH_LONG).show()
                    }
                }
            }
        })
    }


    // -------------------------------------------------------------------------
    // Validation du dossier (Bouton valider)
    // -------------------------------------------------------------------------
    // initValidationButton() -> bouton Valider : vérifie toutes validations puis exporte
    // moveFolderToValid() -> conversion Word -> PDF puis déplacement du dossier
    // convertWordToPdf() -> convertit un Word en PDF via Graph API et supprime l’original
    // moveFolder() -> déplace le dossier validé vers "/Essais/Valide" sur OneDrive
    private fun initValidationButton(context: Context) {
        val buttonValider = findViewById<Button>(R.id.buttonValider)
        buttonValider.setOnClickListener {
            // Vérifie si les trois validations sont à true
            val fpOk = validationJson.optBoolean("FP", false)
            val pveeOk = validationJson.optBoolean("PVEE", false)
            val pveaState = validationJson.optString("PVEA", "non_valide")
            val pveaOk = (pveaState == "validé" || pveaState == "non_necessaire")


            if (fpOk && pveeOk && pveaOk) {
                // ✅ Tous les documents sont validés → afficher la boîte de dialogue
                AlertDialog.Builder(context)
                    .setTitle("Validation du dossier")
                    .setMessage("Êtes-vous sûr de vouloir valider et exporter le dossier $numAffaire ?")
                    .setPositiveButton("Oui") { _, _ ->
                        runOnUiThread {
                            Toast.makeText(context, "Chargement", Toast.LENGTH_SHORT).show()
                        }
                        moveFolderToValid()
                    }
                    .setNegativeButton("Non", null)
                    .show()
            } else {
                // ❌ Il manque au moins une validation → afficher un toast
                Toast.makeText(
                    context,
                    "Impossible de valider : tous les documents ne sont pas validés.",
                    Toast.LENGTH_LONG
                ).show()
            }
        }
    }

    private fun moveFolderToValid() {
        val client = OkHttpClient()

        val listUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$folderId/children"
        val listRequest = Request.Builder()
            .url(listUrl)
            .addHeader("Authorization", "Bearer $authAccessToken")
            .build()

        client.newCall(listRequest).enqueue(object : Callback {
            override fun onFailure(call: Call, e: IOException) {
                runOnUiThread {
                    Toast.makeText(this@DetailActivity, "Erreur lors de la recherche du dossier PV", Toast.LENGTH_LONG).show()
                }
            }

            override fun onResponse(call: Call, response: Response) {
                val body = response.body?.string() ?: ""
                val files = JSONObject(body).optJSONArray("value") ?: JSONArray()
                var validationFileId: String? = null
                var pvFolderId: String? = null

                for (i in 0 until files.length()) {
                    val item = files.getJSONObject(i)
                    val name = item.optString("name")
                    val isFolder = item.has("folder")

                    if (name == "validation.json") {
                        validationFileId = item.optString("id")
                    }

                    if (isFolder && name.equals("PV", true)) {
                        pvFolderId = item.optString("id")
                    }
                }

                if (pvFolderId == null) {
                    runOnUiThread {
                        Toast.makeText(this@DetailActivity, "Aucun dossier PV trouvé", Toast.LENGTH_LONG).show()
                    }
                    return
                }

                // ---- Étape 2 : lister le contenu du dossier PV ----
                val pvUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$pvFolderId/children"
                val pvRequest = Request.Builder()
                    .url(pvUrl)
                    .addHeader("Authorization", "Bearer $authAccessToken")
                    .build()

                client.newCall(pvRequest).enqueue(object : Callback {
                    override fun onFailure(call: Call, e: IOException) {
                        runOnUiThread {
                            Toast.makeText(this@DetailActivity, "Erreur accès dossier PV", Toast.LENGTH_LONG).show()
                        }
                    }

                    override fun onResponse(call: Call, response: Response) {
                        val pvBody = response.body?.string() ?: ""
                        val pvFiles = JSONObject(pvBody).optJSONArray("value") ?: JSONArray()
                        val wordFiles = mutableListOf<JSONObject>()

                        for (i in 0 until pvFiles.length()) {
                            val item = pvFiles.getJSONObject(i)
                            val mime = item.optJSONObject("file")?.optString("mimeType") ?: ""
                            if (mime == "application/vnd.openxmlformats-officedocument.wordprocessingml.document") {
                                wordFiles.add(item)
                            }
                        }

                        // ---- Étape 3 : supprimer validation.json, puis conversion ----
                        val deleteValidation = {
                            if (validationFileId != null) {
                                val deleteUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$validationFileId"
                                val deleteRequest = Request.Builder()
                                    .url(deleteUrl)
                                    .delete()
                                    .addHeader("Authorization", "Bearer $authAccessToken")
                                    .build()

                                client.newCall(deleteRequest).enqueue(object : Callback {
                                    override fun onFailure(call: Call, e: IOException) {
                                        convertWordToPdf(client, wordFiles)
                                    }

                                    override fun onResponse(call: Call, response: Response) {
                                        convertWordToPdf(client, wordFiles)
                                    }
                                })
                            } else {
                                convertWordToPdf(client, wordFiles)
                            }
                        }
                        deleteValidation()
                    }
                })
            }
        })
    }


    private fun convertWordToPdf(
        client: OkHttpClient,
        wordFiles: List<JSONObject>,
        index: Int = 0
    ) {

        if (index >= wordFiles.size) {
            moveFolder(client)
            return
        }

        val wordFile = wordFiles[index]
        val wordId = wordFile.optString("id")
        val fileName = wordFile.optString("name").replace(".docx", ".pdf")

        val url = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$wordId/content?format=pdf"
        val request = Request.Builder()
            .url(url)
            .addHeader("Authorization", "Bearer $authAccessToken")
            .build()

        client.newCall(request).enqueue(object : Callback {
            override fun onFailure(call: Call, e: IOException) {
                runOnUiThread {
                    Toast.makeText(this@DetailActivity, "Erreur conversion $fileName", Toast.LENGTH_LONG).show()
                }
                convertWordToPdf(client, wordFiles, index + 1)
            }

            override fun onResponse(call: Call, response: Response) {
                if (!response.isSuccessful) {
                    runOnUiThread {
                        Toast.makeText(this@DetailActivity, "Erreur conversion $fileName : ${response.code}", Toast.LENGTH_LONG).show()
                    }
                    convertWordToPdf(client, wordFiles, index + 1)
                    return
                }

                val pdfBytes = response.body?.bytes()
                if (pdfBytes != null) {
                    val uploadUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$pvFolderId:/$fileName:/content"
                    val uploadBody = pdfBytes.toRequestBody("application/pdf".toMediaTypeOrNull())

                    val uploadRequest = Request.Builder()
                        .url(uploadUrl)
                        .put(uploadBody)
                        .addHeader("Authorization", "Bearer $authAccessToken")
                        .build()

                    client.newCall(uploadRequest).enqueue(object : Callback {
                        override fun onFailure(call: Call, e: IOException) {
                            runOnUiThread {
                                Toast.makeText(this@DetailActivity, "Erreur upload $fileName : $e", Toast.LENGTH_LONG).show()
                            }
                            convertWordToPdf(client, wordFiles, index + 1)
                        }

                        override fun onResponse(call: Call, response: Response) {

                            // ---- Supprimer le Word original ----
                            val deleteUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$wordId"
                            val deleteRequest = Request.Builder()
                                .url(deleteUrl)
                                .delete()
                                .addHeader("Authorization", "Bearer $authAccessToken")
                                .build()

                            client.newCall(deleteRequest).enqueue(object : Callback {
                                override fun onFailure(call: Call, e: IOException) {
                                    convertWordToPdf(client, wordFiles, index + 1)
                                }

                                override fun onResponse(call: Call, response: Response) {
                                    convertWordToPdf(client, wordFiles, index + 1)
                                }
                            })
                        }
                    })
                } else {
                    convertWordToPdf(client, wordFiles, index + 1)
                }
            }
        })
    }


    private fun moveFolder(client: OkHttpClient) {
        val url = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$folderId"
        val json = JSONObject().apply {
            put("parentReference", JSONObject().put("path", "/1-Essais/2-Valide"))
        }
        val body = json.toString().toRequestBody("application/json".toMediaTypeOrNull())

        val request = Request.Builder()
            .url(url)
            .patch(body)
            .addHeader("Authorization", "Bearer $authAccessToken")
            .build()

        client.newCall(request).enqueue(object : Callback {
            override fun onFailure(call: Call, e: IOException) {
                runOnUiThread {
                    Toast.makeText(this@DetailActivity, "Erreur déplacement", Toast.LENGTH_LONG).show()
                }
            }

            override fun onResponse(call: Call, response: Response) {
                runOnUiThread {
                    if (response.isSuccessful) {
                        Toast.makeText(this@DetailActivity, "Dossier validé", Toast.LENGTH_SHORT).show()

                        // ----- Création du fichier de rappel -----
                        val reminderFileName = "⚠️ Rappel executer script !.txt"
                        val json = JSONObject().apply {
                            put("name", reminderFileName)
                            put("file", JSONObject())
                        }
                        val body = json.toString().toRequestBody("application/json".toMediaTypeOrNull())
                        val createRequest = Request.Builder()
                            .url("https://graph.microsoft.com/v1.0/drives/$driveId/items/$folderId/children")
                            .post(body)
                            .addHeader("Authorization", "Bearer $authAccessToken")
                            .build()

                        client.newCall(createRequest).enqueue(object : Callback {
                            override fun onFailure(call: Call, e: IOException) {
                                runOnUiThread {
                                    Toast.makeText(this@DetailActivity, "Impossible de créer le rappel", Toast.LENGTH_LONG).show()
                                }
                            }

                            override fun onResponse(call: Call, response: Response) {
                                // Optionnel : log ou toast si succès
                            }
                        })
                        // ----------------------------------------

                        val intent = Intent(this@DetailActivity, MainActivity::class.java)
                        intent.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK or Intent.FLAG_ACTIVITY_CLEAR_TASK)
                        startActivity(intent)
                        finish()
                    } else {
                        Toast.makeText(this@DetailActivity, "Erreur déplacement : ${response.code}", Toast.LENGTH_LONG).show()
                    }
                }
            }
        })
    }

}