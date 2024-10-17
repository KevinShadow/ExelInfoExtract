package com.example.demo.controller;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

@Controller
public class XlsController {

    @PostMapping("/upload")
    public String handleFileUpload(@RequestParam("file") MultipartFile file) {
        try {
            // Spécifiez le chemin absolu du répertoire où vous voulez sauvegarder le fichier ZIP
            String uploadDir = System.getProperty("user.dir") + "/temporaire"; // Chemin absolu
            File directory = new File(uploadDir);

            // Créer le répertoire s'il n'existe pas
            if (!directory.exists()) {
                if (directory.mkdirs()) {
                    System.out.println("Répertoire créé : " + uploadDir);
                } else {
                    System.out.println("Échec de la création du répertoire : " + uploadDir);
                }
            }

            // Sauvegarder le fichier ZIP dans le répertoire spécifié
            File zipFile = new File(directory, file.getOriginalFilename());
            file.transferTo(zipFile);

            // Afficher un message de confirmation dans la console
            System.out.println("Fichier ZIP téléchargé : " + zipFile.getAbsolutePath());

            // Extraire les fichiers XLS
            extractXlsFiles(zipFile, directory);

            // Créer le fichier output.xls
            createOutputFile(directory);

            // Lire les fichiers XLS et écrire les données dans output.xls
            writeDataToOutputFile(directory);

            return "redirect:/success"; // Rediriger vers une page de succès (à créer)

        } catch (IOException e) {
            e.printStackTrace();
            return "redirect:/error"; // Rediriger vers une page d'erreur (à créer)
        }
    }

    private void extractXlsFiles(File zipFile, File outputDir) {
        try (ZipInputStream zis = new ZipInputStream(new FileInputStream(zipFile))) {
            ZipEntry zipEntry;
            while ((zipEntry = zis.getNextEntry()) != null) {
                if (zipEntry.getName().endsWith(".xls") || zipEntry.getName().endsWith(".xlsx")) {
                    File newFile = new File(outputDir, zipEntry.getName());

                    // Créer tous les dossiers parents nécessaires
                    new File(newFile.getParent()).mkdirs();

                    // Écrire le fichier
                    try (BufferedOutputStream bos = new BufferedOutputStream(new FileOutputStream(newFile))) {
                        byte[] buffer = new byte[1024];
                        int len;
                        while ((len = zis.read(buffer)) > 0) {
                            bos.write(buffer, 0, len);
                        }
                    }

                    System.out.println("Fichier XLS extrait : " + newFile.getAbsolutePath());
                }
                zis.closeEntry();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void createOutputFile(File outputDir) {
        // Fichier de sortie en format XLS
        File outputFile = new File(outputDir, "output.xls");
        try (HSSFWorkbook workbook = new HSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Données Extraites");

            // Créer un style pour le texte en gras
            CellStyle headerStyle = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true); // Mettre le texte en gras
            headerStyle.setFont(font);

            // Exemple : Ajouter des titres à la première ligne
            Row row = sheet.createRow(0); // Ligne 1
            String[] titre_tout = {"Exercice", "Code Chapitre", "Lib. Chapitre", "Code Programme", "Lib. Programme",
                                   "Code Action", "Lib. Action", "Code Activité", "Lib. Activité", "Code Tâche",
                                   "Lib. Tâche", "Coût Tâche", "Code Nature Tâche", "Fonction Tâche", "Visa Tâche",
                                   "TypeBenef Tâche", "Grandes Masses", "AE Avant-2023", "CP Avant-2023", "AE -2023",
                                   "CP-2023", "AE 2024", "CP 2024", "LR_MN 2024", "AE 2025", "CP 2025", "LR_MN 2025",
                                   "AE 2026", "CP 2026", "LR_MN 2026", "Lib. Nature Tâche", "Est budgétisé ?",
                                   "Version", "TITRE", "TITRE_NBE2", "GRANDE_MASSE_CADRAGE"};

            for (int i = 0; i < titre_tout.length; i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(titre_tout[i]);
                cell.setCellStyle(headerStyle); // Appliquer le style en gras
            }

            // Écriture du fichier
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
                System.out.println("Fichier output.xls créé avec succès avec titres en gras.");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /*private void writeDataToOutputFile(File directory) {
        File outputFile = new File(directory, "output.xls");
        try (Workbook outputWorkbook = new HSSFWorkbook(new FileInputStream(outputFile))) {
            Sheet outputSheet = outputWorkbook.getSheetAt(0); // Feuille déjà créée avec les titres
            int rowCount = outputSheet.getLastRowNum() + 1; // Commencer à la suite de la dernière ligne

            // Parcourir les fichiers XLS extraits
            for (File file : directory.listFiles()) {
                if ((file.getName().endsWith(".xls") || file.getName().endsWith(".xlsx")) && !file.getName().equals("output.xls")) {
                    System.out.println("Tentative d'ouverture du fichier : " + file.getAbsolutePath());
                    try (Workbook workbook = WorkbookFactory.create(file)) {
                        Sheet sheet = workbook.getSheetAt(0); // Lire la première feuille
                        Row row = sheet.getRow(6); // Ligne 7 (index 6)

                        if (row != null) {
                            // Lire les colonnes B, C, D (index 1, 2, 3)
                            String columnB = row.getCell(1) != null ? row.getCell(1).toString() : ""; // Colonne B
                            String columnC = row.getCell(2) != null ? row.getCell(2).toString() : ""; // Colonne C
                            String columnD = row.getCell(3) != null ? row.getCell(3).toString() : ""; // Colonne D

                            // Écrire les données dans le fichier output.xls
                            Row outputRow = outputSheet.createRow(rowCount++);
                            outputRow.createCell(0).setCellValue(columnB); // Colonne B dans output
                            outputRow.createCell(1).setCellValue(columnC); // Colonne C dans output
                            outputRow.createCell(2).setCellValue(columnD); // Colonne D dans output
                        }
                    } catch (IOException e) {
                        System.err.println("Erreur lors de l'ouverture du fichier : " + file.getName() + " - " + e.getMessage());
                    }
                }
            }

            // Écrire le fichier output.xls mis à jour
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                outputWorkbook.write(fos);
                System.out.println("Données écrites dans output.xls : " + outputFile.getAbsolutePath());
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }*/


    private void writeDataToOutputFile(File directory) {
        File outputFile = new File(directory, "output.xls");
        try (Workbook outputWorkbook = new HSSFWorkbook(new FileInputStream(outputFile))) {
            Sheet outputSheet = outputWorkbook.getSheetAt(0); // Feuille déjà créée avec les titres
            int rowCount = outputSheet.getLastRowNum() + 1; // Commencer à la suite de la dernière ligne
    
            // Parcourir les fichiers XLS extraits
            for (File file : directory.listFiles()) {
                if ((file.getName().endsWith(".xls") || file.getName().endsWith(".xlsx")) && !file.getName().equals("output.xls")) {
                    System.out.println("Tentative d'ouverture du fichier : " + file.getAbsolutePath());
                    try (Workbook workbook = WorkbookFactory.create(file)) {
                        Sheet sheet = workbook.getSheetAt(0); // Lire la première feuille
    
                        // Lire les lignes et colonnes nécessaires
                        Row row7 = sheet.getRow(6); // Ligne 7 (index 6)
                        Row row8 = sheet.getRow(7); // Ligne 8 (index 7)
                        Row row10 = sheet.getRow(9); // Ligne 10 (index 9)
    
                        if (row7 != null && row8 != null && row10 != null) {
                            // --- Colonne 2 : B7:D7 (caractères 10 et 11) ---
                            String col2 = row7.getCell(1) != null && row7.getCell(2) != null && row7.getCell(3) != null
                                ? (row7.getCell(1).toString() + row7.getCell(2).toString() + row7.getCell(3).toString()).substring(9, 11)
                                : "";
    
                            // --- Colonne 3 : E7:AH7 ---
                            StringBuilder col3Builder = new StringBuilder();
                            for (int i = 4; i <= 33; i++) { // E7:AH7 correspond aux colonnes 4 à 33 (indexé à partir de 0)
                                if (row7.getCell(i) != null) {
                                    col3Builder.append(row7.getCell(i).toString()).append(" ");
                                }
                            }
                            String col3 = col3Builder.toString().trim();
    
                            // --- Colonne 4 : B8:D8 (caractères 13 et 14) ---
                            String col4 = row8.getCell(1) != null && row8.getCell(2) != null && row8.getCell(3) != null
                                ? (row8.getCell(1).toString() + row8.getCell(2).toString() + row8.getCell(3).toString()).substring(12, 14)
                                : "";
    
                            // --- Colonne 5 : 0 + Colonne 4 + E8:AH8 ---
                            StringBuilder col5Builder = new StringBuilder("0" + col4 + " ");
                            for (int i = 4; i <= 33; i++) { // E8:AH8 correspond aux colonnes 4 à 33 (indexé à partir de 0)
                                if (row8.getCell(i) != null) {
                                    col5Builder.append(row8.getCell(i).toString()).append(" ");
                                }
                            }
                            String col5 = col5Builder.toString().trim();
    
                            // --- Colonne 6 : B10:D10 (caractère 9) ---
                            String col6 = row10.getCell(1) != null && row10.getCell(2) != null && row10.getCell(3) != null
                                ? (row10.getCell(1).toString() + row10.getCell(2).toString() + row10.getCell(3).toString()).substring(8, 9)
                                : "";
    
                            // Écriture des données dans output.xls
                            Row outputRow = outputSheet.createRow(rowCount++);
                            outputRow.createCell(1).setCellValue(col2); // Colonne 2
                            outputRow.createCell(2).setCellValue(col3); // Colonne 3
                            outputRow.createCell(3).setCellValue(col4); // Colonne 4
                            outputRow.createCell(4).setCellValue(col5); // Colonne 5
                            outputRow.createCell(5).setCellValue(col6); // Colonne 6
                        }
                    } catch (IOException e) {
                        System.err.println("Erreur lors de l'ouverture du fichier : " + file.getName() + " - " + e.getMessage());
                    }
                }
            }
    
            // Écriture finale dans le fichier output.xls
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                outputWorkbook.write(fos);
                System.out.println("Données écrites dans output.xls : " + outputFile.getAbsolutePath());
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    
}
