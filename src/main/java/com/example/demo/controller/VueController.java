package com.example.demo.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
//import org.springframework.web.bind.annotation.RequestParam;


@Controller
public class VueController {
    @GetMapping("/")
    public String aceil() {
        return "index";
    }

    @GetMapping("/success")
    public String showSuccessPage() {
        return "success"; // Si la page success.html existe dans templates
    }
    
    @GetMapping("/error")
    public String showErrorPage() {
        return "error"; // Si la page error.html existe dans templates
    }
}
