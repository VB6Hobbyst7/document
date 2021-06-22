import java.util.Scanner;
class Node{
    public int id;
    public Node next;
    public Node(int id){
        this.id = id;
        this.next = null;
    }
}

public class MyLinkedList {
    private Node first;
    public MyLinkedList(){
        first = null;
    }
    public  Node find(int id){
        Node p = first;
        while (p!=null){
            if(p.id == id) return p;
            p = p.next;
        }
        return null;
    }

    public void print(){
//        Node p = first;
//        while (p!=null){
//            System.out.println(p.id +" ");
//            p = p.next;
//        }
        for(Node p = first; p != null; p = p.next ){
            System.out.print(p.id +" ");
        }
        System.out.println();
    }
    public void insertFirst(int id){
        Node i = find(id);
        if(i != null) return;

        Node p = new Node(id);
        p.next = first;
        first = p;
    }

    public void insertLast(int id){

        if(first == null){
            first = new Node(id);
            return;
        }
        Node q = first;
        while (q.next != null) q = q.next;
        Node p = new Node(id);
        q.next = p;
    }
    public Node insertBefore(int u, int v, Node f){

        if(f == null) return null;

        if(f.id == v){
            Node p = new Node(u);
            p.next = f;
            return p;
        }
        f.next = insertBefore(u, v, f.next);
        return f;
    }
    public void insertBefore(int u, int v){
        first = insertBefore(u, v, first);
    }
    public Node insertAfter(int u, int v, Node f){
        if(f == null) return null;
        if(f.id == v){
            Node p = new Node(u);
            p.next = f.next;
            f.next = p;
            return f;
        }
        f.next = insertAfter(u, v, f.next);
        return f;
    }
    public void insertAfter(int u, int v){
        first = insertAfter(u, v, first);
    }
    public Node remove(int u, Node f){
        if(f == null) return f;
        if(f.id == u) return remove(u, f.next);
        f.next = remove(u, f.next);
        return f;
    }
    public void remove(int u){
        first = remove(u, first);
    }
    public static final Scanner input = new Scanner(System.in);
    public static void main(String[] args) {
        int sizeA = input.nextInt();
        MyLinkedList A = new MyLinkedList();
        for(int i = 1; i <= sizeA; i++){
            int item = input.nextInt();
            A.insertLast(item);
        }
        A.print();

        int sizeB = input.nextInt();
        MyLinkedList B = new MyLinkedList();
        for(int i = 1; i <= sizeB; i++){
            int item = input.nextInt();
            B.insertLast(item);
            A.remove(item);
        }
        A.print();
    }
}
