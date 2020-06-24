public class Question3{
	public static void main(String[] args) {
		Question3 q3 = new Question3();
		String str = "abcabcbb";
		int result = q3.lengthOfLongestSubstring(str);
		System.out.println(result);

		System.out.println(q3.hasDuplication("fsdijkyt"));

		
	}


	

	
	public int lengthOfLongestSubstring(String s){
		for(int k=0; k<s.length(); k++){
			for(int i=0, j=s.length()-k; j<=s.length(); i++,j++){
				String sub = s.substring(i, j);
				if(hasDuplication(sub)){
					continue;
				}else{
					return j-i;
				}
			}
		}
		return 0;
	}
	
	


	/*
	//方式一
	public boolean hasDuplication(String str){
		for(int i=0; i<str.length(); i++){
			for(int j=i+1; j<str.length(); j++){
				if(str.charAt(i) == str.charAt(j)){
					return true;
				}
			}
		}

		return false;
	}
	*/

	public boolean hasDuplication(String str){
		StringBuilder sb = new StringBuilder("");
		for(int i=0; i<str.length(); i++){
			char ch = str.charAt(i);
			String s = Character.toString(ch);
			if(sb.toString().contains(s)){
				return true;
			}else{
				sb.append(ch);
			}
			
		}
		return false;
	}

	
}