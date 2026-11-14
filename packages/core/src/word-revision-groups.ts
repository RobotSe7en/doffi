/** Expand logical replacement/move groups once, including linked groups. */
export function expandWordRevisionGroups<T extends {readonly nodeId:string;readonly revisionId?:string|undefined;readonly groupId?:string|undefined;readonly moveId?:string|undefined}>(entries:readonly T[],selectedIds:readonly string[]):readonly T[]{
 const selected=new Set(selectedIds),groups=new Map<string,T[]>(),keys=(entry:T):string[]=>[...entry.groupId?['group:'+entry.groupId]:[],...entry.moveId?['move:'+entry.moveId]:[]];
 for(const entry of entries)for(const key of keys(entry)){const group=groups.get(key)??[];group.push(entry);groups.set(key,group);}
 const queue=entries.filter(entry=>selected.has(entry.revisionId??entry.nodeId)),found=new Set(queue),visited=new Set<string>();
 for(let index=0;index<queue.length;index++)for(const key of keys(queue[index]!)){if(visited.has(key))continue;visited.add(key);for(const entry of groups.get(key)??[])if(!found.has(entry)){found.add(entry);queue.push(entry);}}
 return Object.freeze(entries.filter(entry=>found.has(entry)));
}
