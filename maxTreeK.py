class Solution:
    def maxPathSum(self, root: 'TreeNode') -> 'int':
        def post_order(node):
            cur = node
            stack = []
            res = []
            while cur or stack:
                if cur:
                    stack.append(cur)
                    res.append(cur)
                    cur = cur.right
                else:
                    cur = stack.pop()
                    cur = cur.left
            return res[::-1]
        
        res = -float('inf')
        nodes = post_order(root)
        path_until_node = {None: 0}
        for node in nodes:
            left_max_path = max(0, path_until_node[node.left])
            right_max_path = max(0, path_until_node[node.right])
            res = max(res, left_max_path + right_max_path + node.val)
            path_until_node[node] = max(left_max_path, right_max_path) + node.val
        return res