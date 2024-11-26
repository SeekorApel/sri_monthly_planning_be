package sri.sysint.sri_starter_back.controller;

import java.util.Date;

import javax.persistence.EntityManager;
import javax.persistence.PersistenceContext;
import javax.servlet.http.HttpServletRequest;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import sri.sysint.sri_starter_back.exception.ResourceNotFoundException;
import sri.sysint.sri_starter_back.model.Dummy;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.model.Users;
import sri.sysint.sri_starter_back.repository.DummyRepo;
import sri.sysint.sri_starter_back.repository.UserRepo;

@CrossOrigin(maxAge = 3600)
@RestController
public class SriStarterBackController {
		
	private Response response;	

	@Autowired
	private UserRepo userRepo;
	@Autowired
	private DummyRepo dummyRepo;
	
	@PersistenceContext	
	private EntityManager em;
	
//START - GET MAPPING
	@GetMapping("/sysdate")
	public Response getSysdate(final HttpServletRequest req) throws ResourceNotFoundException {
		Date sys = userRepo.getSysdate();
		response = new Response(new Date(), HttpStatus.OK.value(), null, HttpStatus.OK.getReasonPhrase(),
				req.getRequestURI(), sys);
		return response;
	}
	@GetMapping("/truncSysdate")
	public Response getTruncSysdate(final HttpServletRequest req) throws ResourceNotFoundException {
		Date sys = userRepo.getTruncSysdate();
		response = new Response(new Date(), HttpStatus.OK.value(), null, HttpStatus.OK.getReasonPhrase(),
				req.getRequestURI(), sys);
		return response;
	}
	@GetMapping("/username/{userName}")
	public Response getUserByUserName(final HttpServletRequest req, @PathVariable("userName") String userName) throws ResourceNotFoundException{
		Users users = userRepo.findByUserName(userName);
		if(users == null) {
			response = new Response(new Date(), HttpStatus.NOT_FOUND.value(), null, "DATA NOT FOUND", req.getRequestURI(), users);
		}
		else {
			response = new Response(new Date(), HttpStatus.OK.value(), null, "DATA EXIST", req.getRequestURI(), users);
		}
		return response;
	}	
	@RequestMapping(value = "/insertDataWithGet", method = RequestMethod.GET, headers = "Accept=application/json")
	public Response insertDataDummyWithGet(@RequestParam("code") String code, @RequestParam("desc") String desc, final HttpServletRequest req) {
		Dummy res = dummyRepo.findByCode(code);
		Dummy dt = new Dummy();
		if (res == null) {
			dt.setCode(code);
			dt.setDesc(desc);
			dummyRepo.saveAndFlush(dt);
			response = new Response(new Date(), HttpStatus.OK.value(), null, "INSERTED", req.getRequestURI(), dt);
		}else {
			res.setDesc(desc);
			dummyRepo.saveAndFlush(res);
			response = new Response(new Date(), HttpStatus.OK.value(), null, "UPDATED", req.getRequestURI(), res);
		}
		return response;
	}
//END - GET MAPPING
//START - POST MAPPING
//END - POST MAPPING
//START - PUT MAPPING
//END - PUT MAPPING
//START - DELETE MAPPING
//END - DELETE MAPPING
//START - PROCEDURE
//END - PROCEDURE
}
