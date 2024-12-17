package sri.sysint.sri_starter_back.controller;

import java.util.Date;
import java.util.List;

import javax.servlet.http.HttpServletRequest;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RestController;

import sri.sysint.sri_starter_back.exception.ResourceNotFoundException;
import sri.sysint.sri_starter_back.model.MachineProduct;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.service.MachineProductServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class MachineProductController {
	@Autowired
    private MachineProductServiceImpl machineProductServiceImpl;

    private Response response;
    
    @PostMapping("/saveMachineProduct")
    public Response saveTempMachineProduct(HttpServletRequest req, @RequestBody List<MachineProduct> list) throws ResourceNotFoundException {
       System.out.println("halo");
    	machineProductServiceImpl.deleteAll();  

                for (MachineProduct machineProduct : list) {
                    if (machineProduct.getPART_NUMBER() == null) {
                        throw new IllegalArgumentException("ID_FRONT_REAR is required for all items");
                    }
                }

                List<MachineProduct> saved = machineProductServiceImpl.save(list);
                response = new Response(
                        new Date(),
                        HttpStatus.OK.value(),
                        null,
                        HttpStatus.OK.getReasonPhrase(),
                        req.getRequestURI(),
                        saved
                );
             

        return response;
    }
}
