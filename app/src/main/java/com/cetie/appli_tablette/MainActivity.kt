/*
==============================================================
Script : MainActivity - Menu de sélection d'affaire
Auteur : Deloumeau Maël
Date   : 30/10/2025

But :
- Authentification Microsoft (MSAL) pour accéder à Microsoft Graph
- Scanner ou saisir manuellement un numéro d'affaire
- Rechercher le dossier correspondant sur SharePoint (Graph API)
- Ouvrir un écran de détail avec les informations de l'affaire (DetailActivity)

Librairies principales :
- MSAL : authentification Microsoft (Azure AD)
- OkHttp : requêtes HTTP pour Graph API
- JourneyApps BarcodeScanner : scan de codes-barres
- Android SDK : UI, permissions, Toast, Intent

Étapes principales :
1) Initialisation de l'UI et des boutons (scan + saisie manuelle)
2) Vérification et demande de permission caméra
3) Initialisation du client MSAL et gestion de la connexion
4) Gestion du scan de code-barres (CODE_39)
5) Gestion de la saisie manuelle et validation du numéro
6) Recherche du dossier sur SharePoint via Graph API
7) Lancement de l'écran DetailActivity si le dossier est trouvé
==============================================================
*/
package com.cetie.appli_tablette

// -------------------------------------------------------------------------
// IMPORTS
// -------------------------------------------------------------------------
// Import des librairies Android, MSAL, OkHttp, JSON, scanner de codes-barres

import android.Manifest
import android.content.Intent
import android.content.pm.ActivityInfo
import android.content.pm.PackageManager
import android.os.Bundle
import android.widget.Button
import android.widget.EditText
import android.widget.Toast
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import com.journeyapps.barcodescanner.ScanContract
import com.journeyapps.barcodescanner.ScanOptions
import com.microsoft.identity.client.*
import com.microsoft.identity.client.exception.MsalException
import com.microsoft.identity.client.exception.MsalUiRequiredException
import okhttp3.*
import org.json.JSONObject
import java.io.IOException
import android.view.inputmethod.EditorInfo


// -------------------------------------------------------------------------
// CLASSE PRINCIPALE
// -------------------------------------------------------------------------
// MainActivity : hérite de AppCompatActivity et contient toute la logique
class MainActivity : AppCompatActivity() {

    // -------------------------------------------------------------------------
    // VARIABLES MEMBRES
    // -------------------------------------------------------------------------
    // Permissions, client MSAL, token d'accès
    private val cameraPermission = Manifest.permission.CAMERA
    private val permissonRequestCode = 1234
    private lateinit var pca: ISingleAccountPublicClientApplication
    private var authAccessToken: String? = null

    // -------------------------------------------------------------------------
    // MÉTHODE onCreate
    // -------------------------------------------------------------------------
    // - Initialisation UI (boutons, EditText)
    // - Initialisation MSAL
    // - Vérification permission caméra
    // - Définition des callbacks pour scanner et saisie manuelle

    override fun onCreate(savedInstanceState: Bundle?) {

        val isTablet = resources.configuration.smallestScreenWidthDp >= 600

        requestedOrientation = if (isTablet)
            ActivityInfo.SCREEN_ORIENTATION_LANDSCAPE
        else
            ActivityInfo.SCREEN_ORIENTATION_PORTRAIT

        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        setContentView(R.layout.activity_main)

        val buttonScan = findViewById<Button>(R.id.buttonScan)
        val editNumAffaire = findViewById<EditText>(R.id.editNumAffaire)

        // Initialisation MSAL
        PublicClientApplication.createSingleAccountPublicClientApplication(
            applicationContext,
            R.raw.msal_config,
            object : IPublicClientApplication.ISingleAccountApplicationCreatedListener {
                override fun onCreated(application: ISingleAccountPublicClientApplication) {
                    pca = application
                    checkAccountAndSignIn()
                }
                override fun onError(exception: MsalException) {
                    Toast.makeText(this@MainActivity,"Erreur lors de l'initialisation MSAL: $exception", Toast.LENGTH_LONG).show()
                }
            }
        )

        ensureCameraPermission()

        // SCANNER
        buttonScan.setOnClickListener {
            if (ContextCompat.checkSelfPermission(this, cameraPermission)
                != PackageManager.PERMISSION_GRANTED
            ) {
                ActivityCompat.requestPermissions(
                    this,
                    arrayOf(cameraPermission),
                    permissonRequestCode
                )
            } else {
                startBarcodeScanner()
            }
        }

        // SAISIE MANUELLE
        val buttonEnter = findViewById<Button>(R.id.buttonEnter)
        buttonEnter.setOnClickListener {
            val numero = editNumAffaire.text.toString().trim().uppercase()
            val code39Pattern = Regex("^[A-Z0-9]{8}$")
            if (!code39Pattern.matches(numero)) {
                Toast.makeText(this, "Numéro invalide", Toast.LENGTH_LONG).show()
                return@setOnClickListener
            }
            runOnUiThread {
                Toast.makeText(this@MainActivity, "Chargement", Toast.LENGTH_SHORT).show()
            }
            searchDossierSharePoint(numero)
        }

        editNumAffaire.setOnEditorActionListener { _, actionId, _ ->
            if (actionId == EditorInfo.IME_ACTION_DONE || actionId == EditorInfo.IME_ACTION_GO || actionId == EditorInfo.IME_ACTION_SEND) {
                val numero = editNumAffaire.text.toString().trim().uppercase()
                val code39Pattern = Regex("^[A-Z0-9]{8}$")
                if (!code39Pattern.matches(numero)) {
                    Toast.makeText(this, "Numéro invalide", Toast.LENGTH_LONG).show()
                    return@setOnEditorActionListener true
                }
                Toast.makeText(this, "Chargement", Toast.LENGTH_SHORT).show()
                searchDossierSharePoint(numero)
                true // ✅ indique qu’on a géré l’action
            } else {
                false
            }
        }


    }

    // -------------------------------------------------------------------------
    // RECHERCHE D’UN DOSSIER SUR SHAREPOINT
    // -------------------------------------------------------------------------
    // - Vérifie que l'utilisateur est authentifié
    // - Requête Graph API pour obtenir le site SharePoint
    // - Récupère le drive Documents
    // - Lister le dossier Essais/Temporaire
    // - Cherche le numéro d'affaire et ouvre DetailActivity si trouvé

    private fun searchDossierSharePoint(numero: String) {
        if (authAccessToken.isNullOrEmpty()) {
            Toast.makeText(this, "Vous n'êtes pas connecté, relancez l'application", Toast.LENGTH_LONG).show()
            return
        }

        val client = OkHttpClient()
        val siteUrl = "https://graph.microsoft.com/v1.0/sites/cetie.sharepoint.com:/Sites/Production"

        val siteRequest = Request.Builder()
            .url(siteUrl)
            .addHeader("Authorization", "Bearer $authAccessToken")
            .build()

        client.newCall(siteRequest).enqueue(object : Callback {
            override fun onFailure(call: Call, e: IOException) {
                runOnUiThread {
                    Toast.makeText(this@MainActivity, "Erreur réseau, veuillez vérifier votre connexion Internet : ${e.message}", Toast.LENGTH_LONG).show()
                }
            }

            override fun onResponse(call: Call, response: Response) {
                if (!response.isSuccessful) {
                    runOnUiThread {
                        Toast.makeText(this@MainActivity, "Erreur site HTTP ${response.code}", Toast.LENGTH_LONG).show()
                    }
                    return
                }

                val siteJson = JSONObject(response.body!!.string())
                val siteId = siteJson.getString("id")

                // Récupérer le drive Documents
                val drivesUrl = "https://graph.microsoft.com/v1.0/sites/$siteId/drives"
                val drivesRequest = Request.Builder()
                    .url(drivesUrl)
                    .addHeader("Authorization", "Bearer $authAccessToken")
                    .build()

                client.newCall(drivesRequest).enqueue(object : Callback {
                    override fun onFailure(call: Call, e: IOException) {
                        runOnUiThread {
                            Toast.makeText(this@MainActivity, "Erreur récupération drive : ${e.message}", Toast.LENGTH_LONG).show()
                        }
                    }

                    override fun onResponse(call: Call, response: Response) {
                        if (!response.isSuccessful) {
                            runOnUiThread {
                                Toast.makeText(this@MainActivity, "Erreur drive HTTP ${response.code}", Toast.LENGTH_LONG).show()
                            }
                            return
                        }

                        val drivesJson = JSONObject(response.body!!.string())
                        val drivesArray = drivesJson.getJSONArray("value")
                        if (drivesArray.length() == 0) return

                        val driveId = drivesArray.getJSONObject(0).getString("id")
                        // Lister le dossier Essais/Temporaire
                        val folderPath = "1-Essais/1-Temporaire"
                        val folderUrl = "https://graph.microsoft.com/v1.0/drives/$driveId/root:/$folderPath:/children"
                        val folderRequest = Request.Builder()
                            .url(folderUrl)
                            .addHeader("Authorization", "Bearer $authAccessToken")
                            .build()

                        client.newCall(folderRequest).enqueue(object : Callback {
                            override fun onFailure(call: Call, e: IOException) {
                                runOnUiThread {
                                    Toast.makeText(this@MainActivity, "Erreur récupération dossier PV: ${e.message}", Toast.LENGTH_LONG).show()
                                }
                            }

                            override fun onResponse(call: Call, response: Response) {
                                val body = response.body?.string() ?: ""
                                if (!response.isSuccessful) {
                                    runOnUiThread {
                                        Toast.makeText(this@MainActivity, "Erreur dossier PV HTTP ${response.code}", Toast.LENGTH_LONG).show()
                                    }
                                    return
                                }

                                val filesJson = JSONObject(body).getJSONArray("value")
                                val dossierAffaireJson = (0 until filesJson.length())
                                    .map { filesJson.getJSONObject(it) }
                                    .firstOrNull { it.getString("name").contains(numero) }

                                if (dossierAffaireJson == null) {
                                    runOnUiThread {
                                        Toast.makeText(this@MainActivity, "Affaire non trouvée", Toast.LENGTH_LONG).show()
                                    }
                                    return
                                }


                                val folderName = dossierAffaireJson.getString("name")
                                val folderId = dossierAffaireJson.getString("id")

                                runOnUiThread {
                                    val intent = Intent(this@MainActivity, DetailActivity::class.java)
                                    intent.putExtra("NUM_AFFAIRE", numero)
                                    intent.putExtra("NOM_CLIENT", folderName.split("_").getOrNull(2) ?: "Client")
                                    intent.putExtra("FOLDER_ID", folderId)
                                    intent.putExtra("DRIVE_ID", driveId)
                                    startActivity(intent)
                                }
                            }
                        })
                    }
                })
            }
        })
    }

    // -------------------------------------------------------------------------
    // MSAL : AUTHENTIFICATION MICROSOFT
    // -------------------------------------------------------------------------
    // - Vérifie si un compte est déjà connecté
    // - Sign-in interactif si nécessaire
    // - Acquisition silencieuse d'un token pour Graph API
    // - Gestion des erreurs et MFA
    private fun checkAccountAndSignIn() {
        pca.getCurrentAccountAsync(object : ISingleAccountPublicClientApplication.CurrentAccountCallback {
            override fun onAccountLoaded(activeAccount: IAccount?) {
                if (activeAccount != null) {
                    acquireTokenSilent()
                } else {
                    Toast.makeText(this@MainActivity, "Pas de compte, connexion interactive", Toast.LENGTH_SHORT).show()
                    signIn()
                }
            }

            override fun onAccountChanged(priorAccount: IAccount?, currentAccount: IAccount?) {}
            override fun onError(exception: MsalException) {}
        })
    }

    private fun signIn() {
        runOnUiThread {
            Toast.makeText(this@MainActivity, "Chargement", Toast.LENGTH_SHORT).show()
        }
        pca.signIn(
            this,
            null,
            arrayOf("Files.ReadWrite"),
            object : AuthenticationCallback {
                override fun onSuccess(authenticationResult: IAuthenticationResult) {
                    authAccessToken = authenticationResult.accessToken
                    Toast.makeText(this@MainActivity, "Connexion réussie", Toast.LENGTH_SHORT).show()
                }

                override fun onError(exception: MsalException) {
                    Toast.makeText(this@MainActivity, "Erreur connexion : $exception", Toast.LENGTH_SHORT).show()
                }

                override fun onCancel() {
                    Toast.makeText(this@MainActivity, "Connexion annulée par l'utilisateur", Toast.LENGTH_SHORT).show()
                }
            }
        )
    }

    private fun acquireTokenSilent() {
        val scopes = arrayOf("Files.ReadWrite.All")
        val authority = "https://login.microsoftonline.com/69798bd7-8897-4dfe-9f84-1a861de23af6"

        pca.acquireTokenSilentAsync(scopes, authority, object : AuthenticationCallback {
            override fun onSuccess(result: IAuthenticationResult) {
                authAccessToken = result.accessToken
            }

            override fun onError(exception: MsalException) {
                runOnUiThread {
                    when {
                        exception is MsalUiRequiredException -> {
                            Toast.makeText(
                                this@MainActivity,
                                "Authentification interactive requise (MFA possible)",
                                Toast.LENGTH_SHORT
                            ).show()
                            signIn()
                        }
                        exception.cause is java.net.UnknownHostException ||
                                exception.cause is java.net.ConnectException ||
                                exception.message?.contains("network", true) == true -> {
                            // Cas de réseau indisponible
                            Toast.makeText(
                                this@MainActivity,
                                "Erreur réseau : veuillez vérifier votre connexion Internet",
                                Toast.LENGTH_LONG
                            ).show()
                        }
                        else -> {
                            // Autres erreurs MSAL
                            Toast.makeText(
                                this@MainActivity,
                                "Erreur MSAL : $exception",
                                Toast.LENGTH_LONG
                            ).show()
                        }
                    }
                }
            }
            override fun onCancel() {
                runOnUiThread {
                    Toast.makeText(this@MainActivity, "Authentification annulée", Toast.LENGTH_SHORT).show()
                }
            }
        })
    }

    // -------------------------------------------------------------------------
    // SCANNER DE CODES-BARRES
    // -------------------------------------------------------------------------
    // - Configuration et lancement du scanner CODE_39
    // - Validation du code-barres scanné
    // - Lancement de la recherche du dossier SharePoint de l'affaire
    private fun startBarcodeScanner() {
        val options = ScanOptions()
        options.setDesiredBarcodeFormats(ScanOptions.CODE_39)
        options.setPrompt("Placez le code-barres dans le cadre")
        options.setCameraId(0)
        options.setBeepEnabled(true)
        options.setBarcodeImageEnabled(false)
        barcodeLauncher.launch(options)
    }

    private val barcodeLauncher = registerForActivityResult(ScanContract()) { result ->
        val contents = result.contents
        if (contents == null) {
            Toast.makeText(this, "Scan annulé", Toast.LENGTH_SHORT).show()
            return@registerForActivityResult
        }
        val numAffaire = contents.trim().take(8).uppercase()
        val code39Pattern = Regex("^[A-Z0-9]{8}$")
        if (!code39Pattern.matches(numAffaire)) {
            Toast.makeText(this, "Code invalide (attendu 8 caractères CCode39).", Toast.LENGTH_LONG).show()
            return@registerForActivityResult
        }
        runOnUiThread {
            Toast.makeText(this@MainActivity, "Chargement", Toast.LENGTH_SHORT).show()
        }
        searchDossierSharePoint(numAffaire)
    }

    // -------------------------------------------------------------------------
    // GESTION DE LA PERMISSION CAMERA
    // -------------------------------------------------------------------------
    // - Vérifie si la permission caméra est accordée
    // - Demande de permission si nécessaire
    // - Callback pour gérer le résultat de la demande

    private fun ensureCameraPermission() {
        if (ContextCompat.checkSelfPermission(this, Manifest.permission.CAMERA)
            != PackageManager.PERMISSION_GRANTED) {

            ActivityCompat.requestPermissions(
                this,
                arrayOf(Manifest.permission.CAMERA),
                permissonRequestCode
            )
        }
    }


    override fun onRequestPermissionsResult(requestCode: Int, permissions: Array<out String>, grantResults: IntArray) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults)
        if (requestCode == permissonRequestCode) {
            if (grantResults.isNotEmpty() && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
                Toast.makeText(this, "Permission caméra accordée", Toast.LENGTH_SHORT).show()
            }
            else {
                Toast.makeText(this, "Permission caméra requise pour scanner.", Toast.LENGTH_LONG).show()
            }
        }
    }

}
