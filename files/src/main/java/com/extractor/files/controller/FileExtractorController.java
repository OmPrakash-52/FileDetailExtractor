package com.extractor.files.controller;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


import org.apache.batik.transcoder.TranscoderException;
import org.apache.batik.transcoder.TranscoderInput;
import org.apache.batik.transcoder.TranscoderOutput;
import org.apache.batik.transcoder.image.PNGTranscoder;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.xmlbeans.XmlObject;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequestMapping("/api")
public class FileExtractorController {

	private static final String UPLOAD_DIR = "C:/uploads/";
	
	@PostMapping("/svg")
	public ResponseEntity<Map<String, Object>> uploadSvg(@RequestPart("file") MultipartFile file) {

	    if (!file.getContentType().equals("image/svg+xml")) {
	        return ResponseEntity.status(HttpStatus.UNSUPPORTED_MEDIA_TYPE)
	                .body(Map.of("error", "Only SVG files are allowed!"));
	    }

	    String svgFileName = file.getOriginalFilename();
	    Map<String, Object> result = new HashMap<>();

	    try (InputStream inputStream = file.getInputStream()) {

	        // Parse the SVG file dynamically
	        Document doc = Jsoup.parse(inputStream, "UTF-8", "", org.jsoup.parser.Parser.xmlParser());
	        doc.outputSettings().syntax(Document.OutputSettings.Syntax.xml);

	        Element svgElement = doc.selectFirst("svg");
	        if (svgElement != null && !svgElement.hasAttr("xmlns:xlink")) {
	            svgElement.attr("xmlns:xlink", "http://www.w3.org/1999/xlink");
	        }

	        // Extract all path data dynamically
	        List<String> svgPaths = doc.select("path").eachAttr("d");
	        result.put("svgPaths", svgPaths);

	        // Extract foreignObjects dynamically
	        List<String> foreignObjectsList = doc.select("foreignObject").eachText();
	        result.put("foreignObjects", foreignObjectsList);

	        // Extract links dynamically
	        List<String> linksList = doc.select("a").eachAttr("xlink:href");
	        if (linksList.isEmpty()) {
	            linksList = doc.select("a").eachAttr("href");
	        }
	        result.put("links", linksList);

	        // Extract images dynamically
	        List<String> imagesList = doc.select("image").eachAttr("xlink:href");
	        if (imagesList.isEmpty()) {
	            imagesList = doc.select("image").eachAttr("href");
	        }
	        result.put("images", imagesList);

	        // Convert SVG to PNG dynamically
	        String pngFileName = svgFileName.replace(".svg", ".png");
	        File outputFile = new File(UPLOAD_DIR + pngFileName);

	        try (OutputStream outputStream = new FileOutputStream(outputFile)) {
	            PNGTranscoder transcoder = new PNGTranscoder();
	            TranscoderInput input = new TranscoderInput(new ByteArrayInputStream(doc.html().getBytes(StandardCharsets.UTF_8)));
	            TranscoderOutput output = new TranscoderOutput(outputStream);
	            transcoder.transcode(input, output);
	            result.put("pngFilePath", outputFile.getAbsolutePath());
	        } catch (TranscoderException e) {
	            e.printStackTrace();
	            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
	                    .body(Map.of("error", "Error converting SVG to PNG: " + e.getMessage()));
	        }

	        result.put("message", "SVG parsed and converted successfully");
	        return ResponseEntity.ok(result);

	    } catch (IOException e) {
	        e.printStackTrace();
	        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
	                .body(Map.of("error", "Error processing file: " + e.getMessage()));
	    }
	}


	@PostMapping("/ppt")
	public ResponseEntity<Map<String, Object>> uploadPpt(@RequestPart("file") MultipartFile file) {
		if (!file.getContentType().equals("application/vnd.ms-powerpoint") && !file.getContentType()
				.equals("application/vnd.openxmlformats-officedocument.presentationml.presentation")) {
			return ResponseEntity.status(HttpStatus.UNSUPPORTED_MEDIA_TYPE)
					.body(Map.of("error", "Only PPT or PPTX files are allowed!"));
		}

		Map<String, Object> result = new HashMap<>();
		List<Map<String, Object>> shapesList = new ArrayList<>();

		try (XMLSlideShow ppt = new XMLSlideShow(file.getInputStream())) {
			int slideIndex = 0;
			for (XSLFSlide slide : ppt.getSlides()) {
				for (XSLFShape shape : slide.getShapes()) {
					Map<String, Object> shapeInfo = new HashMap<>();
					shapeInfo.put("slideIndex", slideIndex);

					if (shape instanceof XSLFTextShape) {
						XSLFTextShape textShape = (XSLFTextShape) shape;
						shapeInfo.put("shapeName", textShape.getShapeType());
						shapeInfo.put("text", textShape.getText());
						shapeInfo.put("x", textShape.getAnchor().getX());
						shapeInfo.put("y", textShape.getAnchor().getY());
						shapeInfo.put("width", textShape.getAnchor().getWidth());
						shapeInfo.put("height", textShape.getAnchor().getHeight());

						if (textShape.getShapeType() == ShapeType.ROUND_RECT) {
							extractYellowDot(textShape, shapeInfo);
						}
						System.out.println("Yellow Dot " + shapeInfo.get("yellowDot"));
						shapeInfo.put("svgPath",
								generatePath(textShape.getAnchor().getX(), textShape.getAnchor().getY(),
										textShape.getAnchor().getWidth(), textShape.getAnchor().getHeight(),
										shapeInfo.get("yellowDot")));

					} else if (shape instanceof XSLFAutoShape) {
						XSLFAutoShape autoShape = (XSLFAutoShape) shape;
						ShapeType shapeType = autoShape.getShapeType();
						shapeInfo.put("shapeName", shapeType != null ? shapeType.name() : "Unknown Shape");
						shapeInfo.put("x", autoShape.getAnchor().getX());
						shapeInfo.put("y", autoShape.getAnchor().getY());
						shapeInfo.put("width", autoShape.getAnchor().getWidth());
						shapeInfo.put("height", autoShape.getAnchor().getHeight());

						// For Rounded Rectangle or Shapes with Adjustable Points
						XmlObject[] xmlObjects = autoShape.getXmlObject().selectPath(
								"declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' .//a:adj");

						if (xmlObjects.length > 0) {
							double yellowDot = Double.parseDouble(xmlObjects[0].newCursor().getTextValue());
							shapeInfo.put("yellowDot", yellowDot);
						}
					}
					shapesList.add(shapeInfo);
				}
				slideIndex++;
			}
			result.put("shapes", shapesList);
			result.put("message", "PPT parsed successfully");
			return ResponseEntity.ok(result);
		} catch (IOException e) {
			e.printStackTrace();
			return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
					.body(Map.of("error", "Error processing file: " + e.getMessage()));
		}
	}

	private void extractYellowDot(XSLFShape shape, Map<String, Object> shapeInfo) {
		try {
			XmlObject[] xmlObjects = shape.getXmlObject().selectPath(
					"declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' .//a:gd[@name='adj']");
			if (xmlObjects.length > 0) {
				String yellowDotValue = xmlObjects[0].xmlText();
				if (yellowDotValue.contains("fmla=\"val ")) {
					yellowDotValue = yellowDotValue.split("fmla=\"val ")[1].split("\"")[0];
					if (!yellowDotValue.isEmpty()) {
						double yellowDot = Double.parseDouble(yellowDotValue) / 100000.0; // Convert EMU
						shapeInfo.put("yellowDot", yellowDot);
					}
				}
			}
		} catch (Exception e) {
			shapeInfo.put("yellowDot", "Not Found");
		}
	}

	private String generatePath(double x, double y, double width, double height, Object yellowDotObj) {
		double radius = 0;
		if (yellowDotObj instanceof Double) {
			radius = (Double) yellowDotObj * Math.min(width, height);
		}
		return String.format(
				"M %.2f %.2f Q %.2f %.2f %.2f %.2f L %.2f %.2f Q %.2f %.2f %.2f %.2f L %.2f %.2f Q %.2f %.2f %.2f %.2f L %.2f %.2f Q %.2f %.2f %.2f %.2f Z",
				x + radius, y, x, y, x, y + radius, x, y + height - radius, x, y + height, x + radius, y + height,
				x + width - radius, y + height, x + width, y + height, x + width, y + height - radius, x + width,
				y + radius, x + width, y, x + width - radius, y);
	}
}