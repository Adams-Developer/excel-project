package javatest;

import java.io.FileWriter;
import java.io.IOException;

import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

public class TheChallenge {

	public static void main(String[] args) throws IOException{
		// TODO Auto-generated method stub	
		
		String url = "C:\\challenge";
	
		//create the a json obj, array 
		JSONObject obj = new JSONObject();
		JSONArray multiplyObj = new JSONArray();
		//JSONArray divideObj new JSONArray();
		//JSONArray wordObj new JSONArray();

		
		try {
			// calls the method
		}

		
		finally {
			//if no errors happens code still executes
		}

		
		try {
			

			obj.put("id", "tadams29@mail.depaul.edu");
			
			multiplyObj.put("2400");
			multiplyObj.put("7600");
			multiplyObj.put("150840");
			multiplyObj.put("14355");
			obj.put("numberSetOne", multiplyObj);
			
/*			divideObj.put("8");
			divideObj.put("40");
			divideObj.put("2");
			divideObj.put("2");
			obj.put("numberSetTwo", divideObj);*/
			
			multiplyObj.put("Mess With");
			multiplyObj.put("The Best");
			multiplyObj.put("Die Like");
			multiplyObj.put("The Rest");
			obj.put("wordSetOne", multiplyObj);
			
			
		}
		
		catch(JSONException e) {
			//exception thrown
		}

		//write the file
		try (FileWriter file = new FileWriter("challenge.txt")) {
			file.write(obj.toString());
			System.out.println("JSON object file created");
			System.out.println("\nchallenge: " + obj);
		}
		
		

	}

}
