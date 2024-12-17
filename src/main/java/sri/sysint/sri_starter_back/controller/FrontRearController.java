package sri.sysint.sri_starter_back.controller;

import static sri.sysint.sri_starter_back.security.SecurityConstants.SECRET;

import java.math.BigDecimal;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Optional;

import javax.servlet.http.HttpServletRequest;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.*;

import com.auth0.jwt.JWT;
import com.auth0.jwt.algorithms.Algorithm;
import com.auth0.jwt.exceptions.JWTVerificationException;

import sri.sysint.sri_starter_back.exception.ResourceNotFoundException;
import sri.sysint.sri_starter_back.model.FrontRear;
import sri.sysint.sri_starter_back.model.Response;
import sri.sysint.sri_starter_back.service.FrontRearServiceImpl;

@CrossOrigin(maxAge = 3600)
@RestController
public class FrontRearController {

    @Autowired
    private FrontRearServiceImpl frontRearService;

    private Response response;

    @GetMapping("/getAllFrontRearItems")
    public Response getAllFrontRearItems(HttpServletRequest req) throws ResourceNotFoundException {
        String header = req.getHeader("Authorization");

        if (header == null || !header.startsWith("Bearer ")) {
            throw new ResourceNotFoundException("JWT token not found or maybe not valid");
        }

        String token = header.replace("Bearer ", "");

        try {
            String user = JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
                    .build()
                    .verify(token)
                    .getSubject();

            if (user != null) {
                List<FrontRear> frontRearList = frontRearService.getAllFrontRear();
                response = new Response(
                        new Date(),
                        HttpStatus.OK.value(),
                        null,
                        HttpStatus.OK.getReasonPhrase(),
                        req.getRequestURI(),
                        frontRearList
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

    @GetMapping("/getFrontRearItemById/{id}")
    public Response getFrontRearItemById(HttpServletRequest req, @PathVariable BigDecimal id) throws ResourceNotFoundException {
        String header = req.getHeader("Authorization");

        if (header == null || !header.startsWith("Bearer ")) {
            throw new ResourceNotFoundException("JWT token not found or maybe not valid");
        }

        String token = header.replace("Bearer ", "");

        try {
            String user = JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
                    .build()
                    .verify(token)
                    .getSubject();

            if (user != null) {
                List<FrontRear> frontRear = frontRearService.getFrontRearById(id);
                if (frontRear.isEmpty()) {
                    throw new ResourceNotFoundException("FrontRear with ID " + id + " not found");
                }
                response = new Response(
                        new Date(),
                        HttpStatus.OK.value(),
                        null,
                        HttpStatus.OK.getReasonPhrase(),
                        req.getRequestURI(),
                        frontRear
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

    @PostMapping("/saveFrontRearItems")
    public Response saveFrontRearItems(HttpServletRequest req, @RequestBody List<FrontRear> frontRearList) throws ResourceNotFoundException {
        String header = req.getHeader("Authorization");

        if (header == null || !header.startsWith("Bearer ")) {
            throw new ResourceNotFoundException("JWT token not found or maybe not valid");
        }

        String token = header.replace("Bearer ", "");

        try {
            String user = JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
                    .build()
                    .verify(token)
                    .getSubject();

            if (user != null) {

                // Tambahkan pemanggilan deleteAllFrontRear sebelum menyimpan data baru
                frontRearService.deleteAllFrontRear();  // Menghapus semua data sebelumnya

                // Validasi ID_FRONT_REAR untuk setiap item
                for (FrontRear frontRear : frontRearList) {
                    if (frontRear.getID_FRONT_REAR() == null) {
                        throw new IllegalArgumentException("ID_FRONT_REAR is required for all items");
                    }
                }

                // Simpan data baru
                List<FrontRear> savedFrontRears = frontRearService.saveFrontRearList(frontRearList);
                response = new Response(
                        new Date(),
                        HttpStatus.OK.value(),
                        null,
                        HttpStatus.OK.getReasonPhrase(),
                        req.getRequestURI(),
                        savedFrontRears
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }




    @PostMapping("/updateFrontRearItem")
    public Response updateFrontRearItem(HttpServletRequest req, @RequestBody FrontRear frontRear) throws ResourceNotFoundException {
        String header = req.getHeader("Authorization");

        if (header == null || !header.startsWith("Bearer ")) {
            throw new ResourceNotFoundException("JWT token not found or maybe not valid");
        }

        String token = header.replace("Bearer ", "");

        try {
            String user = JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
                    .build()
                    .verify(token)
                    .getSubject();

            if (user != null) {
                FrontRear updatedFrontRear = frontRearService.updateFrontRear(frontRear);
                response = new Response(
                        new Date(),
                        HttpStatus.OK.value(),
                        null,
                        HttpStatus.OK.getReasonPhrase(),
                        req.getRequestURI(),
                        updatedFrontRear
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

    @GetMapping("/getFrontRearAllMarketingOrders")
    public Response getAllMarketingOrders(HttpServletRequest req) throws ResourceNotFoundException {
        String header = req.getHeader("Authorization");

        if (header == null || !header.startsWith("Bearer ")) {
            throw new ResourceNotFoundException("JWT token not found or maybe not valid");
        }

        String token = header.replace("Bearer ", "");

        try {
            String user = JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
                    .build()
                    .verify(token)
                    .getSubject();
            

            if (user != null) {
                List<Map<String, Object>> marketingOrders = frontRearService.getAllMarketingOrders();
                response = new Response(
                        new Date(),
                        HttpStatus.OK.value(),
                        null,
                        HttpStatus.OK.getReasonPhrase(),
                        req.getRequestURI(),
                        marketingOrders
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

    @GetMapping("/getMarketingOrdersByFrontRearId/{id}")
    public Response getMarketingOrdersByFrontRearId(HttpServletRequest req, @PathVariable BigDecimal id) throws ResourceNotFoundException {
        String header = req.getHeader("Authorization");

        if (header == null || !header.startsWith("Bearer ")) {
            throw new ResourceNotFoundException("JWT token not found or maybe not valid");
        }

        String token = header.replace("Bearer ", "");

        try {
            String user = JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
                    .build()
                    .verify(token)
                    .getSubject();

            if (user != null) {
                List<Map<String, Object>> marketingOrders = frontRearService.getMarketingOrdersByFrontRearId(id);
                response = new Response(
                        new Date(),
                        HttpStatus.OK.value(),
                        null,
                        HttpStatus.OK.getReasonPhrase(),
                        req.getRequestURI(),
                        marketingOrders
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (Exception e) {
            throw new ResourceNotFoundException("JWT token is not valid or expired");
        }

        return response;
    }

    @PostMapping("/deleteAllFrontRearItems")
    public Response deleteAllFrontRearItems(HttpServletRequest req) throws ResourceNotFoundException {
        String header = req.getHeader("Authorization");

        if (header == null || !header.startsWith("Bearer ")) {
            throw new ResourceNotFoundException("JWT token not found or maybe not valid");
        }

        String token = header.replace("Bearer ", "");

        try {
            String user = JWT.require(Algorithm.HMAC512(SECRET.getBytes()))
                    .build()
                    .verify(token)
                    .getSubject();

            if (user != null) {
                frontRearService.deleteAllFrontRear();
                response = new Response(
                        new Date(),
                        HttpStatus.OK.value(),
                        "All FrontRear records deleted",
                        HttpStatus.OK.getReasonPhrase(),
                        req.getRequestURI(),
                        null
                );
            } else {
                throw new ResourceNotFoundException("User not found");
            }
        } catch (JWTVerificationException e) {
            throw new ResourceNotFoundException("Invalid or expired JWT token: " + e.getMessage());
        } catch (Exception e) {
            throw new ResourceNotFoundException("Unexpected error: " + e.getMessage());
        }


        return response;
    }
}
