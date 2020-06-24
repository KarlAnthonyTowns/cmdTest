public class Question2 {

    public static void main(String[] args) {

        ListNode l1 = new ListNode(-1);
        ListNode cur1 = l1;
        ListNode l2 = new ListNode(-1);
        ListNode cur2 = l2;

        for(int i=0;i<5;i++){
            cur1.next = new ListNode(i+3);
            cur1 = cur1.next;
        }

        for(int i=0;i<5;i++){
            cur2.next = new ListNode(i+1);
            cur2 = cur2.next;
        }

        

        l1 = l1.next;
        l2 = l2.next;
        

        ListNode result = new Question2().addTwoNumbers(l1, l2);


        while(l1 != null){
            System.out.print(l1.val);
            l1 = l1.next;
        }
        System.out.println();

        while(l2 != null){
            System.out.print(l2.val);
            l2 = l2.next;
        }
        System.out.println();


        while(result != null){
            System.out.print(result.val);
            result = result.next;
        }

        System.out.println("************");
        
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
        StringBuilder sb = sb.reverse();
        String s = sb.toString();
        /*
        将str转为字符串数组
        */
        char[] ch = s.toCharArray();

        ListNode ln = new ListNode(-1);
        ListNode cur = ln;

        for(int i=0;i<ch.length;i++){
            cur.next = new ListNode(Integer.parstInt(ch[i]));
            cur = cur.next;
        }

        ln = ln.next;

    }


}