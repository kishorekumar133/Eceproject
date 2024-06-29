package com.example.transformdemo;

//public class TranscalcController {
//}
//package com.example.transformdemo.controller;
//package com.example.transformdemo.service;

//package com.example.transformdemo.controller;
//package com.example.transformdemo.service;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;

import java.io.IOException;

@Controller
public class TranscalcController {
    @Autowired
    private TranscalcService transcalcService;

    @GetMapping("/")
    public String index() {
        return "index";
    }

    @PostMapping("/generate")
    public String generateExcel(@RequestParam("numSets") int numSets, Model model) {
        try {
            transcalcService.generateExcel(numSets);
            model.addAttribute("message", "Excel file generated successfully!");
        } catch (IOException e) {
            model.addAttribute("message", "Error generating Excel file: " + e.getMessage());
        }
        return "result";
    }

    @GetMapping("/excel")
    public String excel(){
        return "../../../../GFGsheet.xlsx";
    }
}
