public class Question1{

	public static void main(String[] args){
		int[] nums = {2, 7, 11, 15};
		int target = 20;
		Question1 question1 = new Question1();
		int[] result = question1.twoSum(nums, target);
		for(int i=0;i<result.length;i++){
			System.out.println(result[i]);
		}
		
	}

	public int[] twoSum(int[] nums, int target) {
        int[] result = new int[2];
        for(int i=0;i<nums.length;i++){
            for(int j=i+1;j<+nums.length;j++){
                if(target == nums[i]+nums[j]){
                    result[0] = i;
                    result[1] = j;
                    break;
                }
            }
        }
        return result;
    }
}