using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DogGenUI
	{
	/// <summary>
	///	NodeType represent the type of each node in a Hierarchy,
	/// The values, represent the Hierarchy Level
	/// </summary>
	enum enumNodeTypes
		{
		POR=1,    //Service Portfolio
		FRA=2,	//Framework Portfolio
		FAM=10,	//Service Family
		PRO=20,	//Service Product
		ELE=31,	//Service Element
		FEA=32,	//Service Feature
		ELD=41,	//Element Deliverable
		ELR=42,	//Element Report
		ELM=43,	//Element Meeting
		FED=45,	//Feature Deliverable
		FER=46,	//Feature Report
		FEM=47,	//Feature Meeting
		ESL=51,	//Element Service Level
		FSL=61,	//Feature Service Level
		EAC=71,	//Element Activity
		FAC=72,	//Feature Activity
		}
	class Hierarchy
		{
		private int _sequence;
		public int Sequence
			{
			get
				{
				return this._sequence;
				}
			private set
				{
				this._sequence = value;
				}
			}
		private int _level;
		public int Level
			{
			get
				{
				return this._level;
				}
			private set
				{
				this._level = value;
				}
			}
		private enumNodeTypes _nodeType;
		public enumNodeTypes NodeType
			{
			get
				{
				return this._nodeType;
				}
			private set
				{
				this._nodeType = value;
				}
			}
		private int _nodeID;
		public int NodeID
			{
			get
				{
				return this._nodeID;
				}
			private set
				{
				this._nodeID = value;
				}
			}
		private int _parentNodeID;
		public int ParentNodeID
			{
			get
				{
				return this._parentNodeID;
				}
			private set
				{
				this._parentNodeID = value;
				}
			}
		//private string _nodeTitle;
		//public string NodeTitle
		//	{
		//	get
		//		{
		//		return this._nodeTitle;
		//		}
		//	private set
		//		{
		//		this._nodeTitle = value;
		//		}
		//	}
		//private int _contentLayer1 = 0;
		//public int ContentLayer1
		//	{
		//	get
		//		{
		//		return this._contentLayer1;
		//		}
		//	private set
		//		{
		//		this._contentLayer1 = value;
		//		}
		//	}
		//private int _contentLayer2 = 0;
		//public int ContentLayer2
		//	{
		//	get
		//		{
		//		return this._contentLayer2;
		//		}
		//	private set
		//		{
		//		this._contentLayer2 = value;
		//		}

		//	}
		//private bool _isFrameworkNode;
		//public bool IsFramewrorkNode
		//	{
		//	get
		//		{
		//		return this._isFrameworkNode;
		//		}
		//	private set
		//		{
		//		this._isFrameworkNode = value;
		//		}
		//	}

		/// <summary>
		/// The Construct Hierarchy method reads the string parameter passed as the first parameter and build a List of
		/// Hierarchy Objects based on the parStringNodes content. A List of Hierarchy objects must be passed by REFERENCE as the second paremeter.
		/// </summary>
		/// <param name="parStringNodes"> input string parameter must not be Null</param>
		/// <param name="parHierarchyNodes"> List of Hierarchy objects must be passed by REFERENCE.</param>
		/// <returns>Returns a boolean value which is True of the method completed successfully else it returns false.</returns>
		public static bool ConstructHierarchy(string parStringNodes, ref List<Hierarchy> parHierarchyNodes)
			{
			// Add the statements to build the Hierarchical Structure
			int nodeSeq = 0;
			int nodePosition = 0;
			string thisNodeString;
			int noHierarchyErrors = 0;
			// Process all the nodes in the parSelectedNodes
			do
				{
				// Break if there are no more nodes to process
				if(parStringNodes.IndexOf("<", nodePosition) < 0 | parStringNodes.IndexOf(">", nodePosition) < 0)
					{
					break;
					}
				// Extract a node to process...
				thisNodeString = parStringNodes.Substring(parStringNodes.IndexOf("<", nodePosition) + 1, ((parStringNodes.IndexOf(">", nodePosition) - 1) - parStringNodes.IndexOf("<", nodePosition)));
				// Console.WriteLine("Processing: <{0}>", thisNodeString);
				// Define a new instance of the Hierarchy object
				nodeSeq += 1;
				Hierarchy currentNode = new Hierarchy();
				// Set the Node Sequence
				currentNode.Sequence = nodeSeq;
				// Determine the Node Level
				if(!int.TryParse(thisNodeString.Substring(0, 1), out currentNode._level))
					{
					Console.WriteLine("Level Format error in this node: {0} which is node number {1}.", thisNodeString, nodeSeq);
					noHierarchyErrors += 1;
					}
				// Determine the NodeType
				// if(NodeTypes.IsDefined(typeof(NodeTypes), thisNodeString.Substring(thisNodeString.IndexOf(":") + 1, 3)))
				//Console.WriteLine("NodeType:[{0}]", thisNodeString.Substring(thisNodeString.IndexOf(":") + 1, 3));
				if(Enum.TryParse<enumNodeTypes>(thisNodeString.Substring(thisNodeString.IndexOf(":") + 1, 3), out currentNode._nodeType))
					{
					//Console.WriteLine("NodeType: {0}", currentNode._nodeType);
					}
				else
					{
					Console.WriteLine("NodeType is invalid in this node: {0} which is node number {1}.", thisNodeString, nodeSeq);
					noHierarchyErrors += 1;
					}
				// Determine the Node Identifier
				//Console.WriteLine("NodeID: {0}", thisNodeString.Substring(thisNodeString.IndexOf(";") + 1, (thisNodeString.IndexOf(",")) - (thisNodeString.IndexOf(";") + 1)));
				if(!int.TryParse(thisNodeString.Substring(thisNodeString.IndexOf(";") + 1, (thisNodeString.IndexOf(",")) - (thisNodeString.IndexOf(";") + 1)), out currentNode._nodeID))
					{
					Console.WriteLine("ID is not numeric in this node: {0} which is node number {1}.", thisNodeString, nodeSeq);
					noHierarchyErrors += 1;
					}
				// Determine the Parent Node Identifier
				//Console.WriteLine("ParentNodeID: {0}", thisNodeString.Substring(thisNodeString.IndexOf(",") + 1, (thisNodeString.IndexOf("=")) - (thisNodeString.IndexOf(",") + 1)));
				if(!int.TryParse(thisNodeString.Substring(thisNodeString.IndexOf(",") + 1, (thisNodeString.IndexOf("=")) - (thisNodeString.IndexOf(",") + 1)), out currentNode._parentNodeID))
					{
					Console.WriteLine("Parent ID is not numeric in this node: {0} which is node number {1}.", thisNodeString, nodeSeq);
					noHierarchyErrors += 1;
					}
				//		currentNode._nodeType = thisNodeString.Substring(thisNodeString.IndexOf(":") + 1, 3);

				// All the informarion of the node is gathered now
				// Console.WriteLine("\t\t + Node#:{0} \t\tLevel:{1} Type:{2} ID:{3} ParentID:{4}", currentNode.Sequence, currentNode.Level, currentNode.NodeType, currentNode.NodeID, currentNode.ParentNodeID);
				// Add the object to the List
				parHierarchyNodes.Add(currentNode);
				// Check if there are more nodes to process
				if(parStringNodes.IndexOf(">", nodePosition) > 0)
					{
					nodePosition = parStringNodes.IndexOf(">", nodePosition) + 1;
					}
				else
					{
					nodePosition = parStringNodes.Length;
					}
				}
			while(nodePosition < parStringNodes.Length);
			if(noHierarchyErrors == 0)
				{
				Console.WriteLine("\t No errors occurred and {0} nodes were loaded.", parHierarchyNodes.Count);
				return true;
				}
			else
				{
				Console.WriteLine("{0} errors occurred while loading the nodes.", noHierarchyErrors);
				return false;
				}
			}
			// End of Method
		//End of Class
		}

	}
