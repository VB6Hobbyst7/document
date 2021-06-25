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
-----------------------GFG Traveling Salesman Problem--------------------
import java.util.*;
class GFG{

    static int V = 4;

    // implementation of traveling
// Salesman Problem
    static int travllingSalesmanProblem(int graph[][],
                                        int s)
    {
        // store all vertex apart
        // from source vertex
        ArrayList<Integer> vertex =
                new ArrayList<Integer>();

        for (int i = 0; i < V; i++)
            if (i != s)
                vertex.add(i);

        // store minimum weight
        // Hamiltonian Cycle.
        int min_path = Integer.MAX_VALUE;
        do
        {
            // store current Path weight(cost)
            int current_pathweight = 0;

            // compute current path weight
            int k = s;

            for (int i = 0;
                 i < vertex.size(); i++)
            {
                current_pathweight +=
                        graph[k][vertex.get(i)];
                k = vertex.get(i);
            }
            current_pathweight += graph[k][s];

            // update minimum
            min_path = Math.min(min_path,
                    current_pathweight);

        } while (findNextPermutation(vertex));

        return min_path;
    }

    // Function to swap the data
// present in the left and right indices
    public static ArrayList<Integer> swap(
            ArrayList<Integer> data,
            int left, int right)
    {
        // Swap the data
        int temp = data.get(left);
        data.set(left, data.get(right));
        data.set(right, temp);

        // Return the updated array
        return data;
    }

    // Function to reverse the sub-array
// starting from left to the right
// both inclusive
    public static ArrayList<Integer> reverse(
            ArrayList<Integer> data,
            int left, int right)
    {
        // Reverse the sub-array
        while (left < right)
        {
            int temp = data.get(left);
            data.set(left++,
                    data.get(right));
            data.set(right--, temp);
        }

        // Return the updated array
        return data;
    }

    // Function to find the next permutation
// of the given integer array
    public static boolean findNextPermutation(
            ArrayList<Integer> data)
    {
        // If the given dataset is empty
        // or contains only one element
        // next_permutation is not possible
        if (data.size() <= 1)
            return false;

        int last = data.size() - 2;

        // find the longest non-increasing
        // suffix and find the pivot
        while (last >= 0)
        {
            if (data.get(last) <
                    data.get(last + 1))
            {
                break;
            }
            last--;
        }

        // If there is no increasing pair
        // there is no higher order permutation
        if (last < 0)
            return false;

        int nextGreater = data.size() - 1;

        // Find the rightmost successor
        // to the pivot
        for (int i = data.size() - 1;
             i > last; i--) {
            if (data.get(i) >
                    data.get(last))
            {
                nextGreater = i;
                break;
            }
        }

        // Swap the successor and
        // the pivot
        data = swap(data,
                nextGreater, last);

        // Reverse the suffix
        data = reverse(data, last + 1,
                data.size() - 1);

        // Return true as the
        // next_permutation is done
        return true;
    }

    // Driver Code
    public static void main(String args[])
    {
        // matrix representation of graph
        int graph[][] = {{0, 102, 115, 220},
                {10, 0, 35, 25},
                {15, 35, 0, 30},
                {20, 25, 30, 0}};
        int s = 1;
        System.out.println(
                travllingSalesmanProblem(graph, s));
    }
}

// This code is contributed by adityapande88
