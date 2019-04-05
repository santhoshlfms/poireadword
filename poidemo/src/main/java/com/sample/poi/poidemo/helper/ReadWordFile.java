package com.sample.poi.poidemo.helper;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.stream.Stream;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class ReadWordFile {
	  static String wordfolderPath = "C:\\Users\\sannelson\\Desktop\\word_source"; // location of word files folder
	  static String txtfolderPath = "C:\\Users\\sannelson\\Desktop\\txt_source"; // this is temp folder to hold the word file contains as txt
	  
	  static String word = "apple"; // word to find
	  
	public static void main(String[] args) throws FileNotFoundException, IOException {
		/**
		 * Load the file in a folder 
		 * Read it and see if a word is present 
		 * if present say which file 
		 */
		
		
		  Files.walk(Paths.get(wordfolderPath))
	        .filter(Files::isRegularFile)
	        .forEach((k)->{
				try {
					readFileAsXWPF(k);
				} catch (FileNotFoundException e) {
					e.printStackTrace();
				} catch (IOException e) {
					e.printStackTrace();
				}
			});
	}
	
	/*
	 * Method write contains of word doc to text file so searching is easy 
	
	 * */
	
	static void writeTofile(XWPFWordExtractor totalTextFromWordDocument, String fileName) {
		String txtFileName = fileName.replace(".docx", ".txt");
		String pathToWrite =txtfolderPath+File.separator+txtFileName;
		Path path  = Paths.get(pathToWrite);
		try {
			Files.write(path, totalTextFromWordDocument.getText().getBytes());
					
					// start finding the word in text file
					readFileAndFindWord(path, word);
					
		}catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * read the file as Docx and get text from it
	 * */
	
	static void readFileAsXWPF(Path filePath) throws FileNotFoundException, IOException {
		XWPFDocument docx = new XWPFDocument(new FileInputStream(filePath.toString()));
	    XWPFWordExtractor we = new XWPFWordExtractor(docx);
	  //  System.out.println(we.getText());
	    writeTofile(we, filePath.getFileName().toString());

	}
	/*
	 * Read the text file and print the file containing it.
	 * */
	static void readFileAndFindWord(Path path, String word) {
		try{
			Stream<String> stream = Files.lines(path);
			stream.forEach((line)->{
				//System.out.println(line);
				String found = line.toLowerCase().contains(word) ? "Found in file :" + path  : " ";
				System.out.println(found);
			});
		}catch(Exception e) {
			e.printStackTrace();
		}
		
	}
}
