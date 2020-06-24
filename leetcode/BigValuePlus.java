public class BigValuePlus {

    public static void main(String[] args) {

        BigValuePlus bvp = new BigValuePlus();

        String plus1 = "437698798793247897942234";
        String plus2 = "484234233489790327894789237489728937";

        ListNode l1 = bvp.getListNode(plus1);
        ListNode l2 = bvp.getListNode(plus2);

        ListNode result = bvp.addTwoNumbers(l1, l2);


        StringBuilder sb = new StringBuilder();
        while(result != null){
            sb.append(result.val);
            result = result.next;
        }
        sb = sb.reverse();
        String sum = sb.toString();

        System.out.println(plus1+"+"+plus2+"="+sum);
    }




    public ListNode addTwoNumbers(ListNode l1, ListNode l2) {
        ListNode dummy = new ListNode(-1);
        ListNode cur = dummy;
        int carry = 0;
        while (l1 != null || l2 != null) {
            int d1 = l1 == null ? 0 : l1.val;
            int d2 = l2 == null ? 0 : l2.val;
            int sum = d1 + d2 + carry;
            carry = sum >= 10 ? 1 : 0;
            cur.next = new ListNode(sum % 10);
            cur = cur.next;
            if (l1 != null) l1 = l1.next;
            if (l2 != null) l2 = l2.next;
        }
        if (carry == 1) cur.next = new ListNode(1);
        return dummy.next;
    }




    public ListNode getListNode(String str){
        /*
        将str转为StringBuilder利用其reverse方法反转功能，再转回String
        */
        StringBuilder sb= new StringBuilder(str);
        sb = sb.reverse();
        String s = sb.toString();
        /*
        将str转为字符串数组
        */
        char[] ch = s.toCharArray();

        ListNode ln = new ListNode(-1);
        ListNode cur = ln;

        for(int i=0;i<ch.length;i++){
            cur.next = new ListNode(Integer.parseInt(Character.toString(ch[i])));
            cur = cur.next;
        }
        return ln.next;
    }


}