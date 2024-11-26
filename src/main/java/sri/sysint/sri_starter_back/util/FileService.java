package sri.sysint.sri_starter_back.util;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

@Component
public class FileService {

	private static final String FILE_DIRECTORY = "//server-data/it-doc$/Development/rino/";	

//	private static final String FILE_DIRECTORY = "/PersADM/Audit/";
//	private static final String FILE_DIRECTORY = "/PersADM/EVERIFY/";

	File outputFile = null;

	public String storeFile(List<MultipartFile> files) throws IOException {
		String res = "";
		for (MultipartFile file : files) {
			if (file.isEmpty()) {
				continue;
			}
			byte[] bytes = file.getBytes();
			Path path = Paths.get(FILE_DIRECTORY + file.getOriginalFilename());
			Files.write(path, bytes);
			System.out.println("FILENAME: " + file.getOriginalFilename());
			res = file.getOriginalFilename();
			System.out.println(res);

		}
		return res;
	}
}
